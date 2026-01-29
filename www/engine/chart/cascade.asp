<!--#INCLUDE virtual="/lib/asp/adovbs.asp"-->
<!--#INCLUDE virtual="/sys_include/db.asp"-->
<!--#INCLUDE virtual="/include/json.asp"-->
<!--#INCLUDE virtual="/include/language.asp"-->
<!--#INCLUDE virtual="/include/user.asp"-->
<!--#INCLUDE virtual="/include/input.asp"-->
<!--#INCLUDE virtual="/include/converter.asp"-->
<!--#INCLUDE virtual="/include/table.asp"-->
<!--#INCLUDE virtual="/include/table_market.asp"-->
<!--#INCLUDE virtual="/include/table_report.asp"-->
<!--#INCLUDE virtual="/include/table_property.asp"-->


<%

'Examples:

'http://itswip.realinsite.com/engine/chart/cascade.asp?change=property_0&firstid=R100&secondid=0&period=103
'http://itswip.realinsite.com/engine/chart/cascade.asp?change=property_0&firstid=R101&secondid=0&period=103
'http://itswip.realinsite.com/engine/chart/cascade.asp?change=market_0&firstid=R100&secondid=0&period=103
'http://itswip.realinsite.com/engine/chart/cascade.asp?change=market_0&firstid=R101&secondid=0&period=103


Set Input=new cInput
period=Input.period
change=Input.change
firstid=Input.firstid
secondid=Input.secondid
Set Input=nothing


result=""

Set tr=new cTableReport


select case change

case "property_0", "property_1"

	report=firstid
	product=secondid
	result=PropertySelector_Ajax(period, tr.GetPropertySelectorType(report), product, tr.GetChartID(report), change="property_1")

case "market_0", "market_1"

	report=firstid
	result=MarketSelector_Ajax(period, tr.GetMarketSelectorType(report), tr.GetChartID(report), change="market_1")

end select

Set tr=nothing


Response.AddHeader "X-JSON", result
Response.Write result

%>
