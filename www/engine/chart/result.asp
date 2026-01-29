<!--#INCLUDE virtual="/lib/asp/adovbs.asp"-->
<!--#INCLUDE virtual="/sys_include/globals.asp"-->
<!--#INCLUDE virtual="/sys_include/db.asp"-->
<!--#INCLUDE virtual="/include/language.asp"-->
<!--#INCLUDE virtual="/include/misc.asp"-->
<!--#INCLUDE virtual="/include/user.asp"-->
<!--#INCLUDE virtual="/include/input.asp"-->
<!--#INCLUDE virtual="/include/page.asp"-->
<!--#INCLUDE virtual="/include/converter.asp"-->
<!--#INCLUDE virtual="/include/chart.asp"-->
<!--#INCLUDE virtual="/include/table.asp"-->
<!--#INCLUDE virtual="/include/table_report.asp"-->
<!--#INCLUDE virtual="/include/table_property.asp"-->
<!--#INCLUDE virtual="/include/table_result.asp"-->
<!--#INCLUDE virtual="/include/xml_db_functions.asp"-->


<%

Set Input=new cInput
step=Input.step : if step<>2 then response.redirect "default.asp"
period=Input.period
market=Input.market
report=Input.report
product=Input.product
property=Input.property
format=Input.format
histlen=Input.histlen
Set Input=nothing


'Response.Write "report=" & report & "<br>"
'Response.Write "product=" & product & "<br>"
'Response.Write "property=" & property & "<br>"
'Response.Write "market=" & market & "<br>"
'Response.Write "histlen=" & histlen & "<br>"


' Exception for Office Vacancy Barometer - properties are fixed to DOWNTOWN_CLASS_AA, DOWNTOWN_CLASS_B, SUBURBAN_CLASS_A and SUBURBAN_CLASS_B
' No longer needed
' if report="R500" then property="3,4,5,6"

' Exception for Face Rate Forecast - Office - Midtown Class A --> in this case the only avaliable market is Montreal
'if report="R300" and product=1 and property="17" then market="6"
if report="R300" and property="17" then market="6"

' Same exception for Valuation Parameters - Office - Midtown Class A --> in this case the only avaliable market is Montreal
' Same exception for Valuation Parameters - Multi-Unit Residential - Montreal Suburban Multi-Unit Residential --> in this case the only avaliable market is Montreal
if report="R101" and (property="17" or property="21") then market="6"


' Exception for Altus InSite Investor Outlook / Valuation Parameters - Retail - Tier II Regional Mall --> in this case markets are actually provinces
'if (report="R100" or report="R101") and product=2 and property="8" then market="14,15,16,17,18,19,20"
if (report="R100" or report="R101" or report="R140") and property="8" then market="14,15,16,17,18,19,20"


include_pdf=false
content_area="GetContentArea"
property_ubound=ubound(split(property, ","))

is_standard=(report<>"S100" and report<>"S101" and report<>"S102" and report<>"S200" and report<>"S300" and report<>"S301" and report<>"S302" and report<>"S303" and report<>"S304" and report<>"S305" and report<>"S306" and report<>"S307" and report<>"S308" and report<>"S309" and report<>"S310" and report<>"S311" and report<>"S312" and report<>"S313" and report<>"S314" and report<>"S315" and report<>"T100" and report<>"T101" and report<>"T102" and report<>"R135" and report<>"R136" and report<>"R137" and report<>"R138" and report<>"R139" and report<>"R140" and report<>"X100" and report<>"R134" and report<>"R123")

if Len(property)<1 and is_standard then content_area="GetErrPropertyMissing"
if (product=0 and property_ubound>0) and is_standard and report<>"R400" and report<>"R401" and report<>"R402" and report<>"R403" and report<>"R404" and report<>"R500" and report<>"R124" and report<>"R127" and report<>"R130" and report<>"R134" then content_area="GetErrPropertyTooMany"
if property_ubound>0 and (report="H400" or report="H401" or report="H402") then content_area="GetErrPropertyTooMany"
if Len(property)<1 and (report="H200" or report="H201" or report="H202" or report="H203" or report="H110" or report="H111" or report="H500") then content_area="GetErrPropertyNoData"




' Disabled Valuation Properties for Downtown Office Land as per Kevin and Sandy's Email on Mar 17, 2009
'if report="R101" and property="1" then content_area="GetErrPropertyNoData"




Set t=new cTableReport
groupid=t.GetGroupId(report)
Set t=nothing

Set u = new cUser
if (groupid=0 or groupid=2) then
	if period=g_CURRENT_PERIOD_ID and u.HasAccessToAllMarketsForReport("itsreport_current")=false then content_area="GetErrNoAccess"
	if period<>g_CURRENT_PERIOD_ID and u.HasAccessToAllMarketsForReport("itsreport_snapshot")=false then content_area="GetErrNoAccess"
end if
if (groupid=1) and u.HasAccessToAllMarketsForReport("itsreport_hist")=false then content_area="GetErrNoAccess"
Set u = nothing

if content_area="GetContentArea" then

	Set xmlchart=new cTableXmlchart

	Set t=new cTableResult
	t.CurrPeriodID=period
	Call t.SetReportMarketProperty(report, market, property)
	Call t.SetHistoryLen(histlen)
	data_html=t.Read

	filename=t.ChartObj.Filename
	id_chart=t.ChartObj.UniqueID

	'Response.Write "id_chart =" & id_chart & "<br />"

	data_xml=""
	id_xml=-1
	report_title=t.ChartObj.Title

	if Len(filename)>0 Then
		Set t.TableObj=xmlchart
		data_xml=t.Read
		id_xml=saveXmlInDB(data_xml)
	end if

	Set t = nothing
	Set xmlchart=nothing

	include_pdf=true

end if


Set p=new cPage

Set c = new cConverter
if (groupid=1) then
	p.Subhead=Lang("form_b_title")
else
	p.Subhead=c.NumToStr("LKP01_PERIOD_NAME", period) & " " & Lang("results")
end if
Set c = nothing


if format=1 then
	p.Include_head=false
	p.Include_pdf=false
	html=p.GetDisplay(content_area, false)
	Response.ContentType = "application/pdf"
	Response.BinaryWrite HtmlToPdf(html)
else
	p.Include_pdf=include_pdf
	html=p.GetDisplay(content_area, false)

	Response.Write html
end if

Set p = nothing
%>

<%

Function GetContentArea

	'Response.Write "<p>Step = " & step & " Report = " & report & " Market = " & market & " Product = " & product & " Property = " & property & "</p>"
	'Response.Write "<p align=""center"">id_xml = " & id_xml & "</p>"

	'Response.Write data_xml & "<br>"
	'Response.Write id_xml & "<br>"
	'Response.Write Server.HtmlEncode(data_xml) & "<br>"


	s = ""

	if id_xml<>-1 then

		if Lang("lang")="fr-ca" then
			chart_img=p.Base_its & "/fre/chart/" & filename & "&amp;chartID=" & id_chart & "&amp;xmlID=" & id_xml
		else
			chart_img=p.Base_its & "/eng/chart/" & filename & "&amp;chartID=" & id_chart & "&amp;xmlID=" & id_xml
		end if

		'chart_img="http://www.altusinsite.com/images/login/altusinsitetrans.gif"

		s = s & "<p align=""center""><img src=""" & chart_img & """ alt=""""></p>" & vbNewLine

		if format<>1 then
			s = s & "<p align=""center"">" & vbNewLine
			s = s & "[<a href=""" & chart_img & "&amp;format=2"">" & Lang("image_hires_jpg") & "</a>]&nbsp;&nbsp;" & vbNewLine
			s = s & "[<a href=""" & chart_img & "&amp;format=3"">" & Lang("image_hires_png") & "</a>]" & vbNewLine
			s = s & "</p>" & vbNewLine
			s = s & "<br>" & vbNewLine
		end if
	else
		s = s & "<h3 class=""report_title"">" & report_title & "</h3>"
	end if



	s = s & "<div align=""center"">" & vbNewLine
	s = s & data_html & vbNewLine
'	s = s & "<p><a href=""result2.asp?chartID=" & id_chart & "&amp;property=" & property & """>Click here to display question text and Benchmark Properties for this question</a></p>"
	s = s & "</div>" & vbNewLine


'Response.Write "id_chart =" & id_chart & "<br />"
'Response.Write "property =" & property & "<br />"

s = s & GetContentAreaBechmark(id_chart, property)


	s = s & "<br>" & vbNewLine
	s = s & "<br>" & vbNewLine

	GetContentArea = s
End Function


Function GetContentAreaBechmark(chart_id, property_id)

	'Response.Write "chart_id=" & chart_id & "<br>"
	'Response.Write "property_id=" & property_id & "<br>"

	if len(property_id)<1 then property_id=-1

	select case chart_id
		case 1
			select case property_id
				case 1
					qh_id = 95
				case 9
					qh_id = 130 ' used to be 94
				case 20
					qh_id = 96
				case 3
					qh_id = 105
				case 5
					qh_id = 99
				case 17
					qh_id = 106
				case 6
					qh_id = 112
				case 7
					qh_id = 131
				case 8
					qh_id = 132
				case 12
					qh_id = 139
				case 13
					qh_id = 140
				case 10
					qh_id = 129
				case 15
					qh_id = 145
				case 18
					qh_id = 124
				case 19
					qh_id = 148
				case 21
					qh_id = 152
				case Else
					qh_id = 0
			end select
		case 2
			select case property_id
				case 21
					qh_id = 152
				case 1
					qh_id = 95
				case 4
					qh_id = 94
				case 5
					qh_id = 99
				case 6
					qh_id = 112
				case 3
					qh_id = 105
				case 8
					qh_id = 132
				case 15
					qh_id = 145
				case 12
					qh_id = 139
				case 13
					qh_id = 140
				case 10
					qh_id = 129
				case 9
					qh_id = 130
				case 7
					qh_id = 131
				case 11
					qh_id = 122
				case 17
					qh_id = 106
				case 18
					qh_id = 124
				case 19
					qh_id = 148
				case 20
					qh_id = 96
				case 23
					qh_id = 119
				case Else
					qh_id = 0
			end select
		case 3
			qh_id = 86 ' for some reason this was 130 for Food Anchored Retail Strip and 86 for all other property types
		case 4
			select case property_id
				case 3
					qh_id = 91
				case 4
					qh_id = 117
				case 5
					qh_id = 100
				case 6
					qh_id = 109
				case Else
					qh_id = 0
			end select
		case 5
			select case property_id
				case 3
					qh_id = 92
				case 6
					qh_id = 109
				case Else
					qh_id = 0
			end select
		case 6
			select case property_id
				case 3
					qh_id = 97
				case 5
					qh_id = 103
				case Else
					qh_id = 0
			end select
		case 7
			select case property_id
				case 3
					qh_id = 98
				case 5
					qh_id = 104
				case Else
					qh_id = 0
			end select
		case 8
			select case property_id
				case 17
					qh_id = 108
				'case 4
				'	qh_id = 110 ' D.K. 110 is 122VPOb and apparently this was wrong (as per Bob). Instead it should be 93 which is 22VPO. So on June 10, 2013 this line was disabled.
				case 5
					qh_id = 102
				case 6
					qh_id = 111
				case 12
					qh_id = 144
				case 13
					qh_id = 142
				case Else
					qh_id = 93
			end select
		case 10
			qh_id = 115
		case 11
			qh_id = 125
		case 12
			qh_id = 154
		case 13
			qh_id = 153
		case 14
			qh_id = 156
		case 15
			qh_id = 157
		case 16
			qh_id = 155
		case 17
			qh_id = 160
		case 18
			qh_id = 158
		case 19
			qh_id = 120
		case 22
			qh_id = 128
		case 23
			qh_id = 146
		case 24
			qh_id = 169
		case 25
			qh_id = 89
		case 26
			select case property_id
				case 9
					qh_id = 137
				case 10
					qh_id = 135
				case Else
					qh_id = 0
			end select
		case 27
			qh_id = 141
		case 28
			qh_id = 150
		case 29
			qh_id = 159
		case 30
			qh_id = 114
		case 32
			' DK - modified based on Kevin's request on Mar 16, 2015
			select case property_id
				case 3
					qh_id = 114
				case 5
					qh_id = 114
				case "3,5"
					qh_id = 114
				case Else
					qh_id = 105
			end select
		case 34
			qh_id = 136
		case 35
			qh_id = 143
		case 36
			qh_id = 147
		case 37
			qh_id = 123
		case 38
			qh_id = 162
		case 39
			qh_id = 161
		case 40
			qh_id = 185
		case 41
			qh_id = 164
		case 42
			qh_id = 163
		case 43
			qh_id = 90
		case 44
			if property_id=18 Then
				 qh_id = 121
			Else
				 qh_id = 86
			end If
		case 45
			qh_id = 165
		case 46
			qh_id = 113
		case 47
			qh_id = 134
		case 48
			qh_id = 127
		case 49
			qh_id = 151
		case 50
			qh_id = 149
		case 51
			qh_id = 116
		case 52
			qh_id = 126
		case 53
			qh_id = 167
		case 55
			qh_id = 133
		case 1001
			qh_id = 132
		case Else
			qh_id = 0
	end select

	Set rs = Server.CreateObject("ADODB.Recordset")

	' ADDED BY CRAIG Nov 2014
	question_text_col = "QuestionText"
	if Lang("lang")="fr-ca" then
		question_text_col = "QuestionTextFr"
	end if

	str_sql = "SELECT b.Benchmark, m.Name AS MarketName, q." & question_text_col & " AS QuestionText FROM dbo.dat_Benchmark AS b INNER JOIN dbo.dat_Market AS m ON b.MarketID = m.MarketID RIGHT OUTER JOIN dbo.dat_QuestionHeader AS q ON b.QuestionHeaderID = q.QuestionHeaderID WHERE (q.QuestionHeaderID = " & qh_id & ")"

	If (property_id="-9999") Then
		str_sql = str_sql &" AND 1=0"
	End If

	rs.Open str_sql, its_conn
	s = ""

	'Response.Write "str_sql=" & str_sql & "<br>"

	if not rs.EOF Then
		s = s & "<div style=""width:100%;text-align: center;""><strong>Question:</strong> <br />" & rs.Fields("QuestionText") & "</div><br /><br /><br />" & vbNewLine

		If NOT rs.Fields("Benchmark")="" Then

			s = s & "<table border=""1"" style=""margin-left: auto;margin-right: auto;"">" & vbNewLine
			s = s & "<tr class=""alternate_0""><th colspan=""8"">" & Lang("menu_benchmarkproperties") & "</th></tr>"  & vbNewLine

			Dim market(100), benchmark(100)

			j = 0
			do while not rs.EOF

				'on error resume next

				MarketName = LangField(rs.Fields("MarketName"))

				If NOT rs.Fields("Benchmark")="" Then
					market(j)= MarketName
					benchmark(j)= rs.Fields("Benchmark")
					j = j+1
				End If

				rs.MoveNext

			Loop

			Dim k

			s = s & "<tr class=""alternate_0"">" & vbNewLine
			for k=0 to j-1
				s = s &  "<th>" &  Server.HTMLEncode(market(k)) & "</th>" & vbNewLine
			next
			s = s & "</tr><tr class=""alternate_1"">" & vbNewLine
			for k=0 to j-1
				s = s & "<td>" & Server.HTMLEncode(benchmark(k)) & "</td>" & vbNewLine
			next
			s = s &" </tr> </table>" & vbNewLine

		End If
	End If

	rs.Close()
	Set rs = nothing

    GetContentAreaBechmark= s

End Function



Function GetErrPropertyMissing
	GetErrPropertyMissing=GetInfoBox(Lang("error_a"), Lang("back"))
End Function

Function GetErrPropertyTooMany
	GetErrPropertyTooMany=GetInfoBox(Lang("error_b"), Lang("back"))
End Function

Function GetErrPropertyNoData
	GetErrPropertyNoData=GetInfoBox(Lang("error_d"), Lang("back"))
End Function

Function GetErrNoAccess
	GetErrNoAccess=GetInfoBox(Lang("error_c"), Lang("back"))
End Function

%>
