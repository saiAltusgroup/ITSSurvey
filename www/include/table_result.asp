<%

' DK, Apr 01, 2008 - version 1.0


Function GetQuarterlyFrequency(chart_id, property_arr, strSuffix)

	'Response.Write "chart_id=" & chart_id &", property_arr=" & property_arr & "<br>"

	strSuffix=""
	frequency=1


	select case chart_id

	case 38 ' Calculation of Reversion Value, ...
		frequency=4

	case 22, 23, 24, 25, 26, 28, 30, 32, 34, 43, 45, 46, 47, 48, 49, 50, 52, 53, 55 ' Rental Rate Inflation, ...
		frequency=4

	case 27, 31, 33, 35 ' Rental Rate Inflation, Multiple Unit Residential Valuation Paramaters
		frequency=2

	'case 20, 21 ' Leasehold Outlook, Leasehold Parameters
	'
	'	strSuffix="_Q1"
	'	frequency=4

	case 1, 2 ' Investor Outlook, Valuation Parameters

		if property_arr="21" then ' Montreal Suburban Multiple Unit Residential
			strSuffix="_Q1Q3"
			frequency=2
		end if

		if property_arr="20" then ' Downtown Class AA Office Leasehold
			strSuffix="_Q1"
			frequency=4
		end if
	
		if property_arr="4" then ' Downtown Class B Office
			strSuffix="_Q2"
			frequency=4
		end if

		if property_arr="23" then ' Downtown Class AA Office Ground Lease
			strSuffix="_Q2"
			frequency=4
		end if
		
		if property_arr="1" then ' Downtown Office Land
			strSuffix="_Q2Q4"
			frequency=2
		end if

		if property_arr="5" then ' Suburban Class A Office
			strSuffix="_Q3"
			frequency=4
		end if

		if property_arr="6" then ' Suburban Class B Office
			strSuffix="_Q1"
			frequency=4
		end if

		if property_arr="11" then ' Enclosed Community Mall
			strSuffix="_Q2Q4"
			frequency=2
		end if

		if property_arr="9" then ' Food Anchored Retail Strip
			strSuffix="_Q1Q3"
			frequency=2
		end if

		if property_arr="10" then ' PowerCentre
			strSuffix="_Q1Q3"
			frequency=2
		end if

		if property_arr="18" then ' Single Tenant
			strSuffix="_Q3"
			frequency=4
		end if

		if property_arr="19" then ' Downtown Multiple Unit Resdiential
			strSuffix="_Q3"
			frequency=4
		end if

		if property_arr="8" then ' Tier II Regional Mall
			strSuffix="_Q2Q4"
			frequency=2
		end if

		if (property_arr="12" or property_arr="13") then ' Multi Tenant Industrial or Single Tenant Industrial
			strSuffix="_From2002"
			frequency=1
		end if

	case 4, 5, 6, 7

		if property_arr="6" then ' Suburban Class B Office
			strSuffix="_Q1_From2002"
			frequency=4
		end if

		if property_arr="5" then ' Suburban Class A Office
			strSuffix="_Q3"
			frequency=4
		end if

		if property_arr="4" then ' Downtown Class B Office
			strSuffix="_Q2"
			frequency=4
		end if

		if property_arr="3" then ' Downtown Class AA Office
			strSuffix="_Q2Q4"
			frequency=2
		end if

		if (chart_id=6 or chart_id=7) and property_arr="3" then
			strSuffix="_Q4"
			frequency=4
		end if

	case 8

		if property_arr="5" then ' Suburban Class A Office
			strSuffix="_Q3"
			frequency=4
		end if


	end select


	GetQuarterlyFrequency=frequency

End Function



Class cTableResult

	Public TableObj
	Public ChartObj
	Public CurrPeriodID
	Public HistoryLen

	Private Filter
	Private FilterMarket
	Private FilterProperty
	Private FilterMarketProperty

	Private MarketArr
	Private PropertyArr

	Private DebugFlag


	' Constructor
	Private Sub Class_Initialize()
		'Set TableObj=new cTableListbox
		Set TableObj=new cTableDatagrid
		Set ChartObj=new cChartUnknown
		Call SetReportMarketProperty(0, "", "")
		Call SetHistoryLen(0)
		TableObj.ReadOnly=true
		CurrPeriodID=g_CURRENT_PERIOD_ID
		DebugFlag=false
	End Sub

	' Destructor
	Private Sub Class_Terminate()
		Set TableObj=nothing
		Set ChartObj=nothing
	End Sub


	Public Function Read()

		Set rs = Server.CreateObject("ADODB.Recordset")

		nr_of_cols=0
		str_sql_final=""
		hist_grouping=false

		response_item_tables=Array("", "dat_ResponseItemInt", "dat_ResponseItemFloat", "dat_ResponseItemTxt")

		For col=0 to ubound(ChartObj.ChartIDArr)

			response_item_table=response_item_tables(ChartObj.QuestionTypeIDArr(col))

			str_filter=Filter

			hist_grouping=false

			' See if each column is for a different Property
			if Instr(str_filter, "(-9999)")>0 then
				str_filter=Replace(str_filter, "(-9999)", "(" & ChartObj.PropertyArr(col) & ")")
			end if


			str_sql = "SELECT * FROM dat_Question " & str_filter & " AND ChartID=" & ChartObj.ChartIDArr(col) & " AND ColumnID=" & ChartObj.ColumnIDArr(col) & " AND QuestionTypeID=" & ChartObj.QuestionTypeIDArr(col) ' & " ORDER BY QuestionID"
			if DebugFlag then response.write "<b>str_filter=</b>" & str_filter & "<br>str_sql=" & str_sql & "<br>"

			str_question_list="-1"

			rs.Open str_sql, its_conn
			do while not rs.eof
				str_question_list=str_question_list & "," & rs("QuestionID")
				rs.MoveNext
			loop
			rs.Close()

			if DebugFlag then response.write "<b>str_question_list=</b>" & str_question_list & "<br>"


			strLookupFieldLabel="Name"
			str_sql = "dat_Question AS q JOIN " & response_item_table & " AS ri ON ri.QuestionID=q.QuestionID JOIN dat_Response AS r ON ri.ResponseID=r.ResponseID "
			str_where = "WHERE r.PeriodID IN (" & ChartObj.PeriodIDArr(col) & ") AND ri.QuestionID IN (" & str_question_list & ")" & ChartObj.FilterForValuesArr(col)


			' sql factory

			GroupingType=ChartObj.GroupingTypeArr(col)

			'Response.Write "GroupingType=" & GroupingType & "<br>"

 			select case GroupingType

			case 1, 2, 3, 4, 5

				str_cols = GetColumns(GroupingType-0, "ri.Value", nr_of_cols)
				str_sql = str_sql & "JOIN dat_Market AS l ON q.MarketID=l.MarketID " & str_where
				str_sql = "SELECT l.MarketID AS TemRowID, l.Name AS TempRowLabel, r.PeriodID, " & str_cols & " FROM " & str_sql & " GROUP BY l.MarketID, l.Name, r.PeriodID" ' & " ORDER BY l.MarketID ASC"

				strLookupTable="dat_Market"
				strLookupFieldID="MarketID"
				strFoot=FilterMarket & " ORDER BY MarketID ASC"


			case 21, 22, 23, 24, 25

				str_cols = GetColumns(GroupingType-20, "ri.Value", nr_of_cols)
				str_sql = str_sql & "JOIN dat_Property AS l ON q.PropertyID=l.PropertyID " & str_where
				str_sql = "SELECT l.PropertyID AS TemRowID, l.Name AS TempRowLabel, r.PeriodID, " & str_cols & " FROM " & str_sql & " GROUP BY l.PropertyID, l.Name, r.PeriodID" ' & " ORDER BY l.PropertyID ASC"

				strLookupTable="dat_Property"
				strLookupFieldID="PropertyID"
				strFoot=FilterProperty & " ORDER BY PropertyID ASC"


			case 51, 52, 53, 54, 55

				str_cols = GetColumns(GroupingType-50, "ri.Value", nr_of_cols)
				str_sql = str_sql & "JOIN lkp_ChartType AS l ON 0=l.ChartID " & str_where
				str_sql = "SELECT l.ChartID AS TemRowID, l.Name AS TempRowLabel, r.PeriodID, " & str_cols & " FROM " & str_sql & " GROUP BY l.ChartID, l.Name, r.PeriodID"

				strLookupTable="lkp_ChartType"
				strLookupFieldID="ChartID"
				strFoot=""


			case 81, 82, 83, 84, 85

				str_cols = GetColumns(GroupingType-80, "ri.Value", nr_of_cols)
				str_sql = str_sql & "JOIN vw_MarketProperty AS l ON q.MarketID=l.MarketID AND q.PropertyID=l.PropertyID " & str_where
				str_sql = "SELECT l.MarketPropertyID AS TemRowID, l.Name AS TempRowLabel, r.PeriodID, " & str_cols & " FROM " & str_sql & " GROUP BY l.MarketPropertyID, l.Name, r.PeriodID" ' & " ORDER BY l.MarketID ASC"

				strLookupTable="vw_MarketProperty_ForBarometer"
				strLookupFieldID="MarketPropertyID"
				strFoot=FilterMarketProperty & " ORDER BY MarketPropertyID ASC"


			' 100 and above are used in histroical charts

			case 121, 122, 123, 124, 125

				hist_grouping=true

				str_cols = GetColumns(GroupingType-120, "ri.Value", nr_of_cols)
				str_sql = str_sql & "JOIN vw_Period AS l ON r.PeriodID=l.PeriodID " & str_where
				str_sql = "SELECT l.PeriodID AS TemRowID, l.Name AS TempRowLabel, " & str_cols & " FROM " & str_sql & " GROUP BY l.PeriodID, l.Name"

				'Response.Write PropertyArr & " - "
				'Response.Write ChartObj.ChartIDArr(col) & " - "

				Call GetQuarterlyFrequency(ChartObj.ChartIDArr(col), PropertyArr, strSuffix)

				'strSuffix=""

				strLookupTable="vw_Period_ForHistory" & strSuffix
				strLookupFieldID="PeriodID"
				strFoot="ORDER BY PeriodID DESC"


			case 141, 142, 143, 144, 145

				hist_grouping=true

				str_cols = GetColumns(GroupingType-140, "ri.Value", nr_of_cols)
				str_sql = str_sql & "JOIN vw_MarketPeriod AS l ON q.MarketID=l.MarketID AND r.PeriodID=l.PeriodID " & str_where
				str_sql = "SELECT l.MarketPeriodID AS TemRowID, l.Name AS TempRowLabel, r.PeriodID, " & str_cols & " FROM " & str_sql & " GROUP BY l.MarketPeriodID, l.Name, r.PeriodID" ' & " ORDER BY l.MarketID ASC"
				' Note: SELECT r.PeriodID and GROUP BY r.PeriodID could be omitted if each l.Name record had a unique name

				'Response.Write PropertyArr & " - "
				'Response.Write ChartObj.ChartIDArr(col) & " - "

				Call GetQuarterlyFrequency(ChartObj.ChartIDArr(col), PropertyArr, strSuffix)

				'strSuffix=""

				strLookupTable="vw_MarketPeriod_ForHistory" & strSuffix
				strLookupFieldID="MarketPeriodID"
				strFoot=FilterMarket & " ORDER BY MarketPeriodID DESC"


			case 151, 152, 153, 154, 155

				hist_grouping=true

				str_cols = GetColumns(GroupingType-150, "ri.Value", nr_of_cols)
				str_sql = str_sql & "JOIN vw_PropertyPeriod AS l ON q.PropertyID=l.PropertyID AND r.PeriodID=l.PeriodID " & str_where
				str_sql = "SELECT l.PropertyPeriodID AS TemRowID, l.Name AS TempRowLabel, r.PeriodID, " & str_cols & " FROM " & str_sql & " GROUP BY l.PropertyPeriodID, l.Name, r.PeriodID" ' & " ORDER BY l.PropertyID ASC"
				' Note: SELECT r.PeriodID and GROUP BY r.PeriodID could be omitted if each l.Name record had a unique name

				strLookupTable="vw_PropertyPeriod_ForHistory"
				strLookupFieldID="PropertyPeriodID"
				strFoot=FilterProperty & " ORDER BY PropertyPeriodID DESC"


			case 161, 162, 163, 164, 165

				hist_grouping=true

				str_cols = GetColumns(GroupingType-160, "ri.Value", nr_of_cols)
				str_sql = str_sql & "JOIN vw_MarketPropertyPeriod AS l ON q.MarketID=l.MarketID AND q.PropertyID=l.PropertyID AND r.PeriodID=l.PeriodID " & str_where
				str_sql = "SELECT l.MarketPropertyPeriodID AS TemRowID, l.Name AS TempRowLabel, r.PeriodID, " & str_cols & " FROM " & str_sql & " GROUP BY l.MarketPropertyPeriodID, l.Name, r.PeriodID" ' & " ORDER BY l.MarketID ASC, l.PropertyID ASC"
				' Note: SELECT r.PeriodID and GROUP BY r.PeriodID could be omitted if each l.Name record had a unique name

				strLookupTable="vw_MarketPropertyPeriod_ForHistory"
				strLookupFieldID="MarketPropertyPeriodID"
				strFoot=FilterMarketProperty & " ORDER BY MarketPropertyPeriodID DESC"


			' -------------

			case 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 1010

				str_cols = GetColumns(1, "ri.Value", nr_of_cols)
				str_sql = str_sql & "" & str_where
				str_sql = "SELECT ri.Value AS TemRowID, ri.Value AS TempRowLabel, r.PeriodID, " & str_cols & " FROM " & str_sql & " GROUP BY ri.Value, r.PeriodID" ' & " ORDER BY ri.Value ASC"

				if GroupingType=10 then
					strLookupTable="lkp_ScaleType"
					strLookupFieldID="ScaleID"
					strFoot="ORDER BY ScaleID ASC"

				elseif GroupingType=11 then
					strLookupTable="lkp_ScoreType"
					strLookupFieldID="ScoreID"
					strFoot="ORDER BY ScoreID ASC"

				elseif GroupingType=12 then
					strLookupTable="lkp_TrendType"
					strLookupFieldID="TrendID"
					strFoot="ORDER BY TrendID ASC"

				elseif GroupingType=13 then
					strLookupTable="lkp_SourceType"
					strLookupFieldID="SourceID"
					strFoot="ORDER BY SourceID ASC"

				elseif GroupingType=14 then
					strLookupTable="lkp_FrequencyType"
					strLookupFieldID="FrequencyID"
					strFoot="ORDER BY FrequencyID ASC"

				elseif GroupingType=15 then
					strLookupTable="lkp_InterestType"
					strLookupFieldID="InterestID"
					strFoot="ORDER BY InterestID ASC"

				elseif GroupingType=16 then
					strLookupTable="lkp_LevelType"
					strLookupFieldID="LevelID"
					strFoot="ORDER BY LevelID ASC"

				elseif GroupingType=17 then
					strLookupTable="lkp_CapitalizeType"
					strLookupFieldID="CapitalizeID"
					strFoot="ORDER BY CapitalizeID ASC"

				elseif GroupingType=18 then
					strLookupTable="lkp_SequenceType"
					strLookupFieldID="SequenceID"
					strFoot="ORDER BY SequenceID ASC"

				elseif GroupingType=19 then
					strLookupTable="lkp_GrossRentType"
					strLookupFieldID="GrossRentID"
					strFoot="ORDER BY GrossRentID ASC"

				elseif GroupingType=1010 then
					strLookupTable="lkp_TimePeriodType"
					strLookupFieldID="TimePeriodID"
					strFoot="ORDER BY TimePeriodID ASC"

				else
					UnknownGrouping()

				end if


			' -------------

			case 2000

				str_cols = GetColumns(1, "ri.Value", nr_of_cols)
				str_sql = str_sql & "" & str_where
				str_sql = "SELECT ri.Value AS TemRowID, ri.Value AS TempRowLabel, r.PeriodID, " & str_cols & " FROM " & str_sql & " GROUP BY ri.Value, r.PeriodID" ' & " ORDER BY ri.Value ASC"

				if GroupingType=2000 then
					strLookupTable="lkp_SellingCostsType"
					strLookupFieldID="SellingID"
					strFoot="ORDER BY SellingID ASC"
				end if


			case -1

				str_cols = GetColumns(1, "ri.Industry", nr_of_cols)
				str_sql = "dat_Response AS ri WHERE ri.PeriodID IN (" & ChartObj.PeriodIDArr(col) & ")" & ChartObj.FilterForValuesArr(col)
				str_sql = "SELECT ri.Industry AS TemRowID, ri.Industry AS TempRowLabel, ri.PeriodID, " & str_cols & " FROM " & str_sql & " GROUP BY ri.Industry, ri.PeriodID" ' & " ORDER BY ri.Industry ASC"

				strLookupTable="lkp_IndustryType"
				strLookupFieldID="IndustryID"
				strFoot="ORDER BY IndustryID ASC"


			case else

				UnknownGrouping()

			end select


			'if DebugFlag then response.write "<b>***str_sql=</b>" & str_sql & "<hr>"

			table="table" & col+1
			str_sql_final = str_sql_final & "LEFT OUTER JOIN (" & str_sql & ") AS " & table & " ON table0." & strLookupFieldID & " = " & table & ".TemRowID "

		next


		str_sql="SELECT "
		if hist_grouping and HistoryLen>0 then str_sql=str_sql & "TOP " & HistoryLen & " "
		str_sql=str_sql & "table0." & strLookupFieldID & " AS RowID, table0." & strLookupFieldLabel & " AS RowLabel"
		for i=0 to nr_of_cols-1
			str_sql=str_sql & ", ISNULL(PointY" & i & ", 0) AS PointY" & i
		next
		str_sql_final=str_sql & " FROM " & strLookupTable & " AS table0 " & str_sql_final
		str_sql_final=str_sql_final & " " & strFoot

		if DebugFlag then response.write "<b>str_sql_final=</b>" & str_sql_final & "<hr>"


		if ChartObj.UniqueID=51 Then
			' Exception: report_S313 (Net Effective Discount Period) does not have product type subheader
			s_property=""
		else
			Set p=new cTableProperty
			s_property=p.ReadNames(PropertyArr, 1)
			Set p = nothing
		end if

		'Response.Write "ChartObj.UniqueID=" & ChartObj.UniqueID & "<br>"
		'Response.Write "PropertyArr=" & PropertyArr & "<br>"
		'Response.Write "s_property=" & s_property & "<br>"

		cols_to_add=0
		read_head = ChartObj.InitCallback(TableObj, s_property, cols_to_add)
		'Response.Write "<hr>read_head=" & read_head & "<hr>"
		read_head = TableObj.TableOpen(nr_of_cols+cols_to_add) & read_head


		redim point_array(nr_of_cols-1)

		read_body=""
		rs.Open str_sql_final, its_conn
		do while not rs.eof

			TableObj.RowID=LangField(rs("RowLabel"))

			for i=0 to nr_of_cols-1
				point_array(i)=rs("PointY"&i)
			next

			read_body = read_body & ChartObj.PointCallback(TableObj, point_array)

			rs.MoveNext

		loop
		rs.Close()

		read_foot = TableObj.TableClose
		Read=ChartObj.FinalCallback(TableObj, read_head, read_body, read_foot)

		Set rs=Nothing

	End Function


	Public Function SetHistorylen(history_len)
		HistoryLen=history_len
	End Function


	Public Function SetReportMarketProperty(report, m_arr, p_arr)

		Set c = new cConverter

		curr_period_id=CurrPeriodID
		curr_period_label=c.NumToStr("LKP01_PERIOD_NAME", curr_period_id)

	'Response.Write "report =" & report & "<br />"

		' chart factory

		if report="R101" and p_arr="23" then report="R101-P23"

		
		select case report

		case "X100"
			Set ChartObj=new cChartPie_RespondentCategories


		case "R100"
			Set ChartObj=new cChartPie_ValuationParams


		case "R101-P23"
			' Special case of R101 - "Downtown Class AA Office Ground Lease" (added by DK on Nov 1, 2010)
			Set ChartObj=new cChartBar_ValuationParams_Dynamic
			ChartObj.UniqueID=2
			ChartObj.Title=Lang("report_R101")
			ChartObj.LabelsArr=Array(Lang("chart_item_desc_b1"))
			ChartObj.RespondentsDisplayFormat=1
			ChartObj.Filename="chart_17.asp?ID=1"


		case "R101"

			Set ChartObj=new cChartBar_ValuationParams

			tmp_arr=Split(p_arr, ",")

			' Exceptions
			if p_arr="1" or p_arr="2" or p_arr="14" then ' 1=Downtown Office Land, 2=Suburban Office Land, 14=Industrial Land
				ChartObj.SetFlavourID(1)
			elseif p_arr="15" or p_arr="19" or p_arr="21" then ' 15=Multi Unit Residential, 19=Downtown Multiple Unit Residential, 21=Montreal Midtown Multiple Unit Residential
				ChartObj.SetFlavourID(2)
			'elseif p_arr="9" or p_arr="10" or p_arr="7" or p_arr="9,10,7" then ' 9=Food Anchored Retail Strip, 10=Power Centre, 7=Tier I Regional Mall
			elseif InArray(tmp_arr, "7") or InArray(tmp_arr, "8") or InArray(tmp_arr, "9") or InArray(tmp_arr, "10") or InArray(tmp_arr, "11") or InArray(tmp_arr, "18") then
				' 7=Tier I Regional Mall, 8=Tier II Regional Mall, 9=Food Anchored Retail Strip, 10=Power Centre, 11=Enclosed Community Mall, 18=Single Tenant
				ChartObj.SetFlavourID(3)
			end if

		case "R200"
			Set ChartObj=new cChartBar_FamilyOfRates
			ChartObj.UniqueID=4
			ChartObj.Title=Lang("report_R200")
			ChartObj.UnitForYAxis=Lang("chart_item_dollarpersqft")

		case "R201"
			Set ChartObj=new cChartBar_FamilyOfRates
			ChartObj.UniqueID=5
			ChartObj.Title=Lang("report_R201")
			ChartObj.UnitForYAxis=Lang("chart_item_dollarpersqft")

		case "R202"
			Set ChartObj=new cChartBar_FamilyOfRates
			ChartObj.UniqueID=6
			ChartObj.Title=Lang("report_R202")
			ChartObj.UnitForYAxis=Lang("chart_item_dollarpersqft")

		case "R203"
			Set ChartObj=new cChartBar_FamilyOfRates
			ChartObj.UniqueID=7
			ChartObj.Title=Lang("report_R203")
			ChartObj.UnitForYAxis=Lang("chart_item_percent")


		case "R300"
			Set ChartObj=new cChartBar_FaceRateForecast
			' Exceptions
			if p_arr="15" then ChartObj.OmitCol2And3=true


		case "R400"
			Set ChartObj=new cChartBar_Barometer
			Call ChartObj.SetTypeAndTime(0, 0) ' Product/Market Barometer, Regular

		case "R401"
			Set ChartObj=new cChartBar_Barometer
			Call ChartObj.SetTypeAndTime(1, 0) ' Location Barometer, Regular

		case "R402"
			Set ChartObj=new cChartBar_Barometer
			Call ChartObj.SetTypeAndTime(2, 0) ' Property Type Barometer, Regular

		case "R403"
			Set ChartObj=new cChartBar_Barometer
			Call ChartObj.SetTypeAndTime(1, 1) ' Location Barometer, Rolling Four Quarters
			curr_period_label=Lang("chart_item_rolling4qtrs")

		case "R404"
			Set ChartObj=new cChartBar_Barometer
			Call ChartObj.SetTypeAndTime(2, 1) ' Property Type Barometer, Rolling Four Quarters
			curr_period_label=Lang("chart_item_rolling4qtrs")


		case "R500"
			Set ChartObj=new cChartBar_Barometer
			Call ChartObj.SetCalcId(1)
			Call ChartObj.SetTypeAndTime(0, 0)
			ChartObj.UniqueID=19
			ChartObj.Title=Lang("report_R500") ' Office Vacancy Barometer, Regular


		case "R110"
			Set ChartObj=new cChartPie_ValuationParams
			ChartObj.UniqueID=20
			ChartObj.Title=Lang("report_R110")

'		case "R111"
'			Set ChartObj=new cChartBar_ValuationParams
'			ChartObj.UniqueID=21
'			ChartObj.Title=Lang("report_R111")
'			'ChartObj.PrevPeriodDelta=4
'			ChartObj.SetFlavourID(3)


		case "R120"
			Set ChartObj=new cChartBar_ValuationParams_Dynamic
			ChartObj.UniqueID=22
			ChartObj.Title=Lang("report_R120")
			ChartObj.LabelsArr=Array(Lang("chart_item_scale_g1"), Lang("chart_item_scale_g2"), Lang("chart_item_scale_g3"), Lang("chart_item_scale_g4"), Lang("chart_item_scale_g5"))

		case "R121"
			Set ChartObj=new cChartBar_ValuationParams_Dynamic
			ChartObj.UniqueID=23
			ChartObj.Title=Lang("report_R121")
			ChartObj.LabelsArr=Array(Lang("chart_item_scale_h1"), Lang("chart_item_scale_h2"))

		case "R122"
			Set ChartObj=new cChartBar_ValuationParams_Dynamic
			ChartObj.UniqueID=24
			ChartObj.Title=Lang("report_R122")
			ChartObj.LabelsArr=Array(Lang("chart_item_scale_h3"), Lang("chart_item_scale_h4"), Lang("chart_item_scale_h5"), Lang("chart_item_scale_h6"))
			ChartObj.Units = Lang("chart_item_dollarpersuite")
			ChartObj.DecPlaces = 0

		case "R123"
			Set ChartObj=new cChartBar_ValuationParams_Dynamic
			ChartObj.UniqueID=25
			ChartObj.Title=Lang("report_R123")
			ChartObj.LabelsArr=Array(Lang("L_DOWNTOWN_CLASS_AA_OFFICE"), Lang("L_DOWNTOWN_CLASS_B_OFFICE"), Lang("L_SUBURBAN_CLASS_A_OFFICE"), Lang("L_SUBURBAN_CLASS_B_OFFICE"))
			ChartObj.RespondentsDisplayFormat=1
			m_arr="3,2,11,6,5,7,4,1"
			p_arr="3,4,5,6"

		case "R124"
			Set ChartObj=new cChartBar_ValuationParams_Dynamic
			ChartObj.UniqueID=26
			ChartObj.Title=Lang("report_R124")
			ChartObj.LabelsArr=Array(Lang("chart_item_scale_i1"), Lang("chart_item_scale_i2"))
			ChartObj.RespondentsDisplayFormat=1

		case "R141"
			Set ChartObj=new cChartBar_ValuationParams_Dynamic
			ChartObj.UniqueID=55
			ChartObj.Title=Lang("report_R124")
			ChartObj.LabelsArr=Array(Lang("chart_item_scale_i1"), Lang("chart_item_scale_i2"))
			ChartObj.RespondentsDisplayFormat=1
			m_arr="14,15,16,17,18,19,20"

		case "R125"
			Set ChartObj=new cChartBar_ValuationParams_Dynamic
			ChartObj.UniqueID=27
			ChartObj.Title=Lang("report_R125")
			ChartObj.LabelsArr=Array(Lang("chart_item_scale_j1"), Lang("chart_item_scale_j2"))
			ChartObj.RespondentsDisplayFormat=1
			ChartObj.Units = Lang("chart_item_dollarpersqft")

		case "R126"
			Set ChartObj=new cChartBar_ValuationParams_Dynamic
			ChartObj.UniqueID=28
			ChartObj.Title=Lang("report_R126")
			ChartObj.LabelsArr=Array(Lang("chart_item_head_f1"))
			ChartObj.RespondentsDisplayFormat=1

'		case "R140"
'			Set ChartObj=new cChartBar_ValuationParams_Dynamic
'			ChartObj.UniqueID=54
'			ChartObj.Title=Lang("report_R140")
'			ChartObj.LabelsArr=Array(Lang("chart_item_head_f5"))
'			ChartObj.RespondentsDisplayFormat=1

		case "R127"
			Set ChartObj=new cChartBar_ValuationParams_Dynamic
			ChartObj.UniqueID=30
			ChartObj.Title=Lang("report_R127")
			ChartObj.LabelsArr=Array(Lang("chart_item_desc_b1"), Lang("chart_item_desc_b2"), Lang("chart_item_desc_b3"), Lang("chart_item_desc_b4"), Lang("chart_item_dollarpersqft"), Lang("chart_max_loan_to_value"), Lang("chart_min_spread"))
			ChartObj.RespondentsDisplayFormat=2
			ChartObj.GroupingBase=20

'		case "R128"
'			Set ChartObj=new cChartBar_ValuationParams_Dynamic
'			ChartObj.UniqueID=31
'			ChartObj.Title=Lang("report_R128")
'			ChartObj.LabelsArr=Array(Lang("chart_item_desc_b1"), Lang("chart_item_desc_b2"), Lang("chart_item_desc_b3"), Lang("avg_dollar_ps"))
'			ChartObj.RespondentsDisplayFormat=3
'			ChartObj.GroupingBase=20

'		case "R129"
'			Set ChartObj=new cChartPie_ValuationParams
'			ChartObj.UniqueID=33
'			ChartObj.Title=Lang("report_R129")

		case "R130"
			Set ChartObj=new cChartPie_ValuationParams
			ChartObj.UniqueID=32
			ChartObj.Title=Lang("report_R130")

		case "R131"
			Set ChartObj=new cChartBar_ValuationParams_Dynamic
			ChartObj.UniqueID=34
			ChartObj.Title=Lang("report_R131")
			ChartObj.LabelsArr=Array(Lang("chart_rent_inflation"), Lang("chart_expense_inflation"))

		case "R132"
			Set ChartObj=new cChartBar_ValuationParams_Dynamic
			ChartObj.UniqueID=35
			ChartObj.Title=Lang("report_R132")
			ChartObj.LabelsArr=Array(Lang("chart_initial_net_rental_rate"))
			ChartObj.Units = Lang("chart_item_dollarpersqft")
			ChartObj.RespondentsDisplayFormat=1
			ChartObj.DecPlaces = 2

		case "R133"
			Set ChartObj=new cChartBar_ValuationParams_Dynamic
			ChartObj.UniqueID=36
			ChartObj.Title=Lang("report_R133")
			ChartObj.LabelsArr=Array(Lang("chart_development_yield"), Lang("chart_stabilized_cap_rate"))
			ChartObj.RespondentsDisplayFormat=1
			ChartObj.DecPlaces = 2

		case "R134"
			Set ChartObj=new cChartBar_DynamicPie
			ChartObj.UniqueID=37
			ChartObj.Title=Lang("report_R134")
			ChartObj.ColumnsArr=Array(1, 2, 3, 4)
			ChartObj.LabelsArr=Array(Lang("L_TIER_I_REGIONAL_MALL"), Lang("L_COMMUNITY_MALL"), Lang("L_POWER_CENTRE"), Lang("L_FOOD_ANCHORED_RETAIL_STRIP"))
			ChartObj.ValuesArr=Array(Lang("L_LEVEL_1"), Lang("L_LEVEL_2"), Lang("L_LEVEL_3"), Lang("L_LEVEL_4"), Lang("L_LEVEL_5"))
			p_arr="11,9,10,7"

		case "R135"
			Set ChartObj=new cChartBar_ValuationParams_Dynamic
			ChartObj.UniqueID=43
			ChartObj.Title=Lang("report_R135")
			ChartObj.LabelsArr=Array(Lang("L_DOWNTOWN_CLASS_AA_OFFICE"), Lang("L_DOWNTOWN_CLASS_B_OFFICE"), Lang("L_SUBURBAN_CLASS_A_OFFICE"), Lang("L_SUBURBAN_CLASS_B_OFFICE"))
			ChartObj.Units = Lang("chart_item_months")
			ChartObj.RespondentsDisplayFormat=1
			ChartObj.DecPlaces = 1
			p_arr = "3,4,5,6"
			m_arr = "3,2,11,6,5,7,4,1"

		case "R136"
			Set ChartObj=new cChartBar_ValuationParams_Dynamic
			ChartObj.UniqueID=45
			ChartObj.Title=Lang("report_R136")
			ChartObj.LabelsArr=Array(Lang("L_DOWNTOWN_CLASS_AA_OFFICE"), Lang("L_DOWNTOWN_CLASS_B_OFFICE"), Lang("L_SUBURBAN_CLASS_A_OFFICE"), Lang("L_TIER_I_REGIONAL_MALL"), Lang("L_MULTI_TENANT_INDUSTRIAL"), Lang("L_SUBURBAN_MULTI_UNIT_RESIDENTIAL"))
			ChartObj.Units = Lang("chart_item_months")
			ChartObj.RespondentsDisplayFormat=1
			ChartObj.DecPlaces = 1
			p_arr = "3,4,12,5,15,7"
			m_arr = "3,2,11,6,5,7,4,1"

		case "R137"
			Set ChartObj=new cChartBar_ValuationParams_Dynamic
			ChartObj.UniqueID=48
			ChartObj.Title=Lang("report_R137")
			ChartObj.LabelsArr=Array(Lang("L_TIER_I_REGIONAL_MALL"), Lang("L_TIER_II_REGIONAL_MALL"), Lang("L_COMMUNITY_MALL"), Lang("L_POWER_CENTRE"), Lang("L_FOOD_ANCHORED_RETAIL_STRIP"))
			ChartObj.RespondentsDisplayFormat=0
			ChartObj.DecPlaces = 1
			p_arr = "11,9,10,7,8"
			m_arr = "3,2,11,6,5,7,4,1"

		case "R138"
			Set ChartObj=new cChartBar_ValuationParams_Dynamic
			ChartObj.UniqueID=49
			ChartObj.Title=Lang("report_R138")
			ChartObj.LabelsArr=Array(Lang("L_SUBURBAN_MULTI_UNIT_RESIDENTIAL"))
			ChartObj.RespondentsDisplayFormat=1
			ChartObj.DecPlaces = 0
			p_arr = "15"
			m_arr = "3,2,11,6,5,7,4,1"

		case "R139"
			Set ChartObj=new cChartBar_ValuationParams_Dynamic
			ChartObj.UniqueID=50
			ChartObj.Title=Lang("report_R139")
			ChartObj.LabelsArr=Array(Lang("L_SUBURBAN_MULTI_UNIT_RESIDENTIAL"))
			ChartObj.RespondentsDisplayFormat=1
			ChartObj.Units = Lang("chart_item_dollarpersuite")
			ChartObj.DecPlaces = 0
			p_arr = "15"
			m_arr = "3,2,11,6,5,7,4,1"

        case "R140"
        	Set ChartObj=new cChartBar_ValuationParams_Dynamic
			ChartObj.UniqueID=55
			ChartObj.Title=Lang("report_R140")
			ChartObj.LabelsArr=Array(Lang("chart_rent_inflation"), Lang("chart_expense_inflation"))

		case "S100"
			Set ChartObj=new cChartBar_Motives
			ChartObj.UniqueID=10
			ChartObj.Title=Lang("report_S100")
			ChartObj.LabelsArr=Array(Lang("chart_item_scale_a1"), Lang("chart_item_scale_a2"), Lang("chart_item_scale_a3"), Lang("chart_item_scale_a4"), Lang("chart_item_scale_a5"), Lang("chart_item_scale_a6"), Lang("chart_item_scale_a7"), Lang("chart_item_scale_a8"), Lang("chart_item_scale_a9"))

		case "S101"
			Set ChartObj=new cChartBar_Motives
			ChartObj.UniqueID=11
			ChartObj.Title=Lang("report_S101")
			ChartObj.LabelsArr=Array(Lang("chart_item_scale_b1"), Lang("chart_item_scale_b2"), Lang("chart_item_scale_b3"), Lang("chart_item_scale_b4"), Lang("chart_item_scale_b5"), Lang("chart_item_scale_b6"), Lang("chart_item_scale_b7"))

		case "S102"
			Set ChartObj=new cChartBar_Motives
			ChartObj.UniqueID=12
			ChartObj.Title=Lang("report_S102")
			ChartObj.LabelsArr=Array(Lang("chart_item_scale_b1"), Lang("chart_item_scale_b2"), Lang("chart_item_scale_b3"), Lang("chart_item_scale_b4"), Lang("chart_item_scale_b5"), Lang("chart_item_scale_b6"), Lang("chart_item_scale_b7"), Lang("chart_item_scale_b8"))

		case "S200"
			Set ChartObj=new cChartBar_Motives
			ChartObj.UniqueID=13
			ChartObj.Title=Lang("report_S200")
			ChartObj.LabelsArr=Array(Lang("chart_item_scale_c1"), Lang("chart_item_scale_c2"), Lang("chart_item_scale_c3"), Lang("chart_item_scale_c4"), Lang("chart_item_scale_c5"), Lang("chart_item_scale_c6"), Lang("chart_item_scale_c7"), Lang("chart_item_scale_c8"))


		case "S300"
			Set ChartObj=new cChartBar_CSCMFamilyA
			ChartObj.UniqueID=14
			ChartObj.SetFlavourID(2)
			ChartObj.Title=Lang("report_S300")
			ChartObj.Subtitle="(Reits and Corporations)"
			ChartObj.LabelsArr=Array(Lang("chart_item_scale_d1"), Lang("chart_item_scale_d2"), Lang("chart_item_scale_d3"))

		case "S301"
			Set ChartObj=new cChartBar_CSCMFamilyB
			ChartObj.UniqueID=15
			ChartObj.Title=Lang("report_S301")
			ChartObj.LabelsArr=Array(Lang("chart_item_scale_d1"), Lang("chart_item_scale_d2"), Lang("chart_item_scale_d3"))

		case "S302"
			Set ChartObj=new cChartBar_CSCMFamilyB
			ChartObj.UniqueID=16
			ChartObj.Title=Lang("report_S302")
			ChartObj.LabelsArr=Array(Lang("chart_item_scale_e1"), Lang("chart_item_scale_e2"))

		case "S303"
			Set ChartObj=new cChartBar_CSCMFamilyA
			ChartObj.SetFlavourID(1)
			ChartObj.UniqueID=17
			ChartObj.Title=Lang("report_S303")
			ChartObj.Subtitle=Lang("report_S303_subtitle")
			ChartObj.LabelsArr=Array(Lang("chart_item_scale_f1"), Lang("chart_item_scale_f2"), Lang("chart_item_scale_f3"), Lang("chart_item_scale_f4"))

		case "S304"
			Set ChartObj=new cChartBar_CSCMFamilyA
			ChartObj.SetFlavourID(0)
			ChartObj.UniqueID=18
			ChartObj.Title=Lang("report_S304")
			ChartObj.Subtitle=Lang("report_S304_subtitle")
			ChartObj.LabelsArr=Array("", "", "")

		case "S305"
			Set ChartObj=new cChartBar_OCR_IRR
			ChartObj.Title=Lang("report_S305")
			ChartObj.ColumnID=1
			m_arr="1,2,3,4,5,6,7,11"
			p_arr="-9999"

		case "S306"
			Set ChartObj=new cChartBar_OCR_IRR
			ChartObj.Title=Lang("report_S306")
			ChartObj.ColumnID=2
			m_arr="1,2,3,4,5,6,7,11"
			p_arr="-9999"

		case "S307"
			Set ChartObj=new cChartPie_ValuationParams
			ChartObj.UniqueID=38
			ChartObj.Title=Lang("report_S307")
			ChartObj.GroupingType=14
			p_arr="-1"

		case "S308"
			Set ChartObj=new cChartPie_ValuationParams
			ChartObj.UniqueID=39
			ChartObj.Title=Lang("report_S308")
			ChartObj.GroupingType=15
			p_arr = "-1"

		case "S309"
			Set ChartObj=new cChartPie_ValuationParams
			ChartObj.UniqueID=40
			ChartObj.Title=Lang("report_S309")
			ChartObj.GroupingType=17
			p_arr="-1"

		case "S310"
			Set ChartObj=new cChartPie_ValuationParams
			ChartObj.UniqueID=44
			ChartObj.Title=Lang("report_S310")
			ChartObj.GroupingType=19
			p_arr="18"

		case "S311"
			Set ChartObj=new cChartBar_ValuationParams_Dynamic
			ChartObj.UniqueID=46
			ChartObj.Title=Lang("report_S311")
			ChartObj.LabelsArr=Array(Lang("L_LT5"), Lang("L_5_10"), Lang("L_10_20"), Lang("L_20_50"), Lang("L_50_100"), Lang("L_GT100"))
			ChartObj.Units = Lang("chart_item_pctgrossincome")
			ChartObj.RespondentsDisplayFormat=1
			ChartObj.DecPlaces = 1
			'ChartObj.CustomSubtitle="All Office"
			m_arr="0"
			p_arr="22" ' All office i.e. 3,4,5,6. Used to be 0.

		case "S312"
			Set ChartObj=new cChartBar_ValuationParams_Dynamic
			ChartObj.UniqueID=47
			ChartObj.Title=Lang("report_S312")
			ChartObj.LabelsArr=Array(Lang("L_TIER_I_REGIONAL_MALL"), Lang("L_TIER_II_REGIONAL_MALL"), Lang("L_COMMUNITY_MALL"), Lang("L_POWER_CENTRE"), Lang("L_FOOD_ANCHORED_RETAIL_STRIP"))
			ChartObj.RespondentsDisplayFormat=0
			ChartObj.DecPlaces = 1
			m_arr="0"
			p_arr="7,8,9,10,11"

		case "S313"
			Set ChartObj=new cChartPie_ValuationParams
			ChartObj.UniqueID=51
			ChartObj.Title=Lang("report_S313")
			ChartObj.GroupingType = 1010

		case "S314"
			Set ChartObj=new cChartBar_ValuationParams_Dynamic
			ChartObj.UniqueID=52
			ChartObj.Title=Lang("report_S314")
			ChartObj.LabelsArr=Array(Lang("L_TIER_I_REGIONAL_MALL"), Lang("L_TIER_II_REGIONAL_MALL"), Lang("L_COMMUNITY_MALL"), Lang("L_POWER_CENTRE"), Lang("L_FOOD_ANCHORED_RETAIL_STRIP"))
			ChartObj.Units = Lang("chart_item_months")
			ChartObj.RespondentsDisplayFormat=0
			ChartObj.DecPlaces = 1
			m_arr="0"
			p_arr="7,8,9,10,11"

		case "S315"
			Set ChartObj=new cChartBar_ValuationParams_Dynamic
			ChartObj.UniqueID=53
			ChartObj.Title=Lang("report_S315")
			ChartObj.LabelsArr=Array(Lang("L_CLASS_A_OFFICE"), Lang("L_CLASS_B_OFFICE"), Lang("L_RETAIL"), Lang("L_INDUSTRIAL"), Lang("L_MULTI_RES"))
			ChartObj.RespondentsDisplayFormat=0
			ChartObj.DecPlaces = 1
			m_arr="0"
			p_arr="5,6,9,12,15"


		case "T100"
			Set ChartObj=new cChartTable_CDMParameters_Dynamic
			ChartObj.UniqueID=29
			ChartObj.Title=Lang("report_T100")
			ChartObj.RespondentsGroupingType=21
			ChartObj.RespondentsQuestionType=2
			ChartObj.RespondentsDisplayFormat=1
			ChartObj.DecPlaces=0
			p_arr="3,12,5,15,7"

			If (CurrPeriodID >= 108) Then

				' Modifed by DK on Sep 11, 2014 as per Kevin's info: 83CSDM - Conventional Debt Market Parameters has a column that is no longer part of the question; The data is inconsistent and/or bad. We don't ask the question part anymore so there won't be any new data.

				'ChartObj.LabelsArr=Array(Lang("chart_max_loan_to_value_ratio"), Lang("chart_minimum_debt"), Lang("chart_longest_amortization"), Lang("chart_min_spread_5yr"), Lang("chart_min_spread_10yr"), Lang("chart_current_availability") & " - " & Lang("L_POOR"),Lang("chart_current_availability") & " - " & Lang("L_GOOD"),Lang("chart_current_availability") & " - " & Lang("L_EXCELLENT"),Lang("chart_current_availability") & " - " & Lang("L_GOOD") & "/" & Lang("L_EXCELLENT"), Lang("chart_preferred_term"), Lang("chart_additional_bps"))
				'ChartObj.ColumnsArr=Array(1, 2, 3, 4, 5, 7,7,7,7, 6, 8)
				'ChartObj.TypesArr=Array(2, 2, 1, 1, 2, 2,2,2,2, 2, 2)
				'ChartObj.GroupingArr=Array(22, 22, 22, 22, 22, 21,21,21,21, 22, 22)
				'ChartObj.FilterArr=Array("", "", "", "", "", " AND ri.Value=1.0"," AND ri.Value=2.0"," AND ri.Value=3.0","", "", "")
				''ChartObj.FilterArr=Array("", "", "", "", "", " "," "," AND ri.Value=350","", "", "")

				ChartObj.LabelsArr=Array(Lang("chart_max_loan_to_value_ratio"), Lang("chart_longest_amortization"), Lang("chart_min_spread_5yr"), Lang("chart_min_spread_10yr"), Lang("chart_current_availability") & " - " & Lang("L_POOR"),Lang("chart_current_availability") & " - " & Lang("L_GOOD"),Lang("chart_current_availability") & " - " & Lang("L_EXCELLENT"),Lang("chart_current_availability") & " - " & Lang("L_GOOD") & "/" & Lang("L_EXCELLENT"), Lang("chart_preferred_term"), Lang("chart_additional_bps"))
				ChartObj.ColumnsArr=Array(1, 3, 4, 5, 7,7,7,7, 6, 8)
				ChartObj.TypesArr=Array(2, 1, 1, 2, 2,2,2,2, 2, 2)
				ChartObj.GroupingArr=Array(22, 22, 22, 22, 21,21,21,21, 22, 22)
				ChartObj.FilterArr=Array("", "", "", "", " AND ri.Value=1.0"," AND ri.Value=2.0"," AND ri.Value=3.0","", "", "")
				'ChartObj.FilterArr=Array("", "", "", "", " "," "," AND ri.Value=350","", "", "")

				Else

				ChartObj.LabelsArr=Array(Lang("chart_max_loan_to_value_ratio"), Lang("chart_min_spread_10yr"), Lang("chart_longest_amortization"), Lang("chart_current_availability") & " - " & Lang("L_POOR"),Lang("chart_current_availability") & " - " & Lang("L_GOOD"),Lang("chart_current_availability") & " - " & Lang("L_EXCELLENT"),Lang("chart_current_availability") & " - " & Lang("L_GOOD") & "/" & Lang("L_EXCELLENT"))
				ChartObj.ColumnsArr=Array(1, 2, 3, 4,4,4,4)
				ChartObj.TypesArr=Array(2, 2, 1, 1,1,1,1)
				ChartObj.GroupingArr=Array(22, 22, 22, 21,21,21,21)
				ChartObj.FilterArr=Array("", "", "", " AND ri.Value=1"," AND ri.Value=2"," AND ri.Value=3","")

			End If

		case "T101"
			Set ChartObj=new cChartTable_CDMParameters_Dynamic
			ChartObj.UniqueID=41
			ChartObj.Title=Lang("report_T101")
			ChartObj.LabelsArr=Array(Lang("chart_dcfirr"), Lang("chart_noi_yield_year"), Lang("chart_noi_yield_avg"), Lang("chart_cf_yield_year"), Lang("chart_cf_yield_avg"), Lang("chart_item_dollarpersqft"))
			ChartObj.ColumnsArr=Array(1, 2, 3, 4, 5, 6)
			ChartObj.TypesArr=Array(1, 1, 1, 1, 1, 1)
			ChartObj.GroupingArr=Array(18, 18, 18, 18, 18, 18)
			ChartObj.FilterArr=Array("", "", "", "", "", "")
			ChartObj.RespondentsGroupingType=18
			ChartObj.RespondentsQuestionType=1
			ChartObj.RespondentsDisplayFormat=0
			ChartObj.DecPlaces=0
			p_arr="3"

'		case "T102"
'			Set ChartObj=new cChartTable_CDMParameters_Dynamic
'			ChartObj.UniqueID=42
'			ChartObj.Title=Lang("report_T102")
'			ChartObj.LabelsArr=Array(Lang("L_5_10"), Lang("L_10_20"), Lang("L_20_50"), Lang("L_50_100"), Lang("L_100_PLUS"))
'			ChartObj.ColumnsArr=Array(1, 2, 3, 4, 5)
'			ChartObj.TypesArr=Array(2, 2, 2, 2, 2)
'			ChartObj.GroupingArr=Array(2000, 2000, 2000, 2000, 2000)
'			ChartObj.FilterArr=Array("", "", "", "", "")
'			ChartObj.RespondentsGroupingType=2000
'			ChartObj.RespondentsQuestionType=1
'			ChartObj.RespondentsDisplayFormat=0
'			ChartObj.DecPlaces=1




		case "H100"
			Set ChartObj=new cChartArea_HistoricalValuationParams
			curr_period_id=-1
			curr_period_label=Lang("chart_item_historical")

			' This is a quick hack becasue it works only if m_arr is a single number. It does not work when m_arr is a comma separated array.
			curr_period_label=curr_period_label & " - " & LangField(c.NumToStr("LKP03_MARKET_NAME", m_arr))

		case "H101", "R102"

			Set ChartObj=new cChartLine_HistoricalValuationParams
			curr_period_id=-1
			curr_period_label=Lang("chart_item_historical")

			tmp_arr=Split(p_arr, ",")

			' Exceptions
			if p_arr="23" then ' 23=Downtown Class AA Office Ground Lease
				ChartObj.SetFlavourID(4)
			elseif p_arr="1" or p_arr="2" or p_arr="14" then ' 1=Downtown Office Land, 2=Suburban Office Land, 14=Industrial Land
				ChartObj.SetFlavourID(1)
			elseif p_arr="15" or p_arr="19" or p_arr="21" then ' 15=Multi Unit Residential, 19=Downtown Multiple Unit Residential, 21=Montreal Midtown Multiple Unit Residential
				ChartObj.SetFlavourID(2)
			''elseif p_arr="9" or p_arr="10" or p_arr="7" or p_arr="9,10,7" then ' 9=Food Anchored Retail Strip, 10=Power Centre, 7=Tier I Regional Mall
			elseif InArray(tmp_arr, "7") or InArray(tmp_arr, "8") or InArray(tmp_arr, "9") or InArray(tmp_arr, "10") or InArray(tmp_arr, "11") or InArray(tmp_arr, "18") then
				' 7=Tier I Regional Mall, 8=Tier II Regional Mall, 9=Food Anchored Retail Strip, 10=Power Centre, 11=Enclosed Community Mall, 18=Single Tenant
				ChartObj.SetFlavourID(3)
			end if

		case "H200"
			Set ChartObj=new cChartLine_HistoricalFamilyOfRates
			ChartObj.UniqueID=4
			ChartObj.Title=Lang("report_H200")
			ChartObj.UnitForYAxis=Lang("chart_item_dollarpersqft")
			curr_period_id=-1
			curr_period_label=Lang("chart_item_historical")

		case "H201"
			Set ChartObj=new cChartLine_HistoricalFamilyOfRates
			ChartObj.UniqueID=5
			ChartObj.Title=Lang("report_H201")
			ChartObj.UnitForYAxis=Lang("chart_item_dollarpersqft")
			curr_period_id=-1
			curr_period_label=Lang("chart_item_historical")

		case "H202"
			Set ChartObj=new cChartLine_HistoricalFamilyOfRates
			ChartObj.UniqueID=6
			ChartObj.Title=Lang("report_H202")
			ChartObj.UnitForYAxis=Lang("chart_item_dollarpersqft")
			curr_period_id=-1
			curr_period_label=Lang("chart_item_historical")

		case "H203"
			Set ChartObj=new cChartLine_HistoricalFamilyOfRates
			ChartObj.UniqueID=7
			ChartObj.Title=Lang("report_H203")
			ChartObj.UnitForYAxis=Lang("chart_item_percent")
			curr_period_id=-1
			curr_period_label=Lang("chart_item_historical")

		case "H400"
			Set ChartObj=new cChartBar_HistoricalBarometer
			Call ChartObj.SetType(0) ' Product/Market Barometer
			curr_period_id=-1
			curr_period_label=Lang("chart_item_historical")

		case "H401"
			Set ChartObj=new cChartBar_HistoricalBarometer
			Call ChartObj.SetType(1) ' Location Barometer
			curr_period_id=-1
			curr_period_label=Lang("chart_item_historical")

		case "H402"
			Set ChartObj=new cChartBar_HistoricalBarometer
			Call ChartObj.SetType(2) ' Property Type Barometer
			curr_period_id=-1
			curr_period_label=Lang("chart_item_historical")

		case "H500"
			Set ChartObj=new cChartBar_HistoricalBarometer
			Call ChartObj.SetCalcId(1)
			Call ChartObj.SetType(0)
			ChartObj.UniqueID=19
			ChartObj.Title=Lang("report_H500") ' Office Vacancy Barometer
			curr_period_id=-1
			curr_period_label=Lang("chart_item_historical")

		case "H110"
			Set ChartObj=new cChartArea_HistoricalValuationParams
			ChartObj.UniqueID=20
			ChartObj.Title=Lang("report_R110")
			curr_period_id=-1
			curr_period_label=Lang("chart_item_historical")

'		case "H111"
'			Set ChartObj=new cChartLine_HistoricalValuationParams
'			ChartObj.UniqueID=21
'			ChartObj.Title=Lang("report_R111")
'			ChartObj.SetFlavourID(3)
'			curr_period_id=-1
'			curr_period_label=Lang("chart_item_historical")

		case else
			Set ChartObj=new cChartUnknown

		end select


		Set c = nothing



		ChartObj.CurrPeriodID=curr_period_id
		ChartObj.CurrPeriodLabel=curr_period_label

		ChartObj.PrevPeriodDelta=GetQuarterlyFrequency(ChartObj.UniqueID, p_arr, strSuffix)
		'Response.Write "ChartObj.UniqueID=" & ChartObj.UniqueID & ", p_arr=" & p_arr & ", ChartObj.PrevPeriodDelta=" & ChartObj.PrevPeriodDelta & "<br>"

		ChartObj.Init


		MarketArr=m_arr
		PropertyArr=p_arr


		FilterMarket=""
		FilterProperty=""

		if Len(m_arr)>0 Then FilterMarket="(MarketID IS NULL OR MarketID IN (" & MarketArr & "))"
		if Len(p_arr)>0 Then FilterProperty="(PropertyID IS NULL OR PropertyID IN (" & PropertyArr & "))"

		s=""

		if len(FilterMarket)>0 and len(FilterProperty)>0 then s=" AND "
		FilterMarketProperty=FilterMarket & s & FilterProperty

		'Filter="WHERE Status=1 AND ChartID=" & ChartObj.UniqueID
		Filter="WHERE Status<>0"

		if len(FilterMarket)>0 then FilterMarket="WHERE " & FilterMarket
		if len(FilterProperty)>0 then FilterProperty="WHERE " & FilterProperty
		if len(FilterMarketProperty)>0 then
			Filter=Filter & " AND " & FilterMarketProperty
			FilterMarketProperty="WHERE " & FilterMarketProperty
		end if

		'Response.Write "ChartObj.UniqueID=" & ChartObj.UniqueID & "<br>"

	End Function


	Private Function GetColumns(id, col_name, nr_of_cols)

		' Note: If AVG is applied to a series of int values, the result will be int. To force SQL Server to interpret int values as float we add "0.0". It's a hack but it seems to work.  
	
		GetColumns=""

		point_y0="PointY" & nr_of_cols : nr_of_cols=nr_of_cols+1

		' cases 3 (MAX) and 4 (MIN) are not presently used

		if id=1 then
			GetColumns="COUNT(" & col_name & ") AS " & point_y0

		elseif id=2 then
			GetColumns="AVG(" & col_name & "+0.0) AS " & point_y0

		elseif id=3 then
			GetColumns="MAX(" & col_name & ") AS " & point_y0

		elseif id=4 then
			GetColumns="MIN(" & col_name & ") AS " & point_y0

		elseif id=5 then
			point_y1="PointY" & nr_of_cols : nr_of_cols=nr_of_cols+1
			point_y2="PointY" & nr_of_cols : nr_of_cols=nr_of_cols+1
			GetColumns="MAX(" & col_name & ") AS " & point_y0 & ", MIN(" & col_name & ") AS " & point_y1 & ", AVG(" & col_name & "+0.0) AS " & point_y2

		else
			UnknownGrouping()

		end if

	End Function


	Private Function UnknownGrouping

		Response.Write "*** UNKNOWN GROUPING TYPE ***"
		Response.End

	End Function

End Class

%>
