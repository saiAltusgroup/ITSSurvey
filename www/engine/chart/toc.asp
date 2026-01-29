<!--#INCLUDE virtual="/lib/asp/adovbs.asp"-->
<!--#INCLUDE virtual="/sys_include/globals.asp"-->
<!--#INCLUDE virtual="/sys_include/db.asp"-->
<!--#INCLUDE virtual="/include/language.asp"-->
<!--#INCLUDE virtual="/include/misc.asp"-->
<!--#INCLUDE virtual="/include/user.asp"-->
<!--#INCLUDE virtual="/include/input.asp"-->
<!--#INCLUDE virtual="/include/page.asp"-->
<!--#INCLUDE virtual="/include/converter.asp"-->
<!--#INCLUDE virtual="/include/table.asp"-->
<!--#INCLUDE virtual="/include/table_period.asp"-->
<!--#INCLUDE virtual="/include/table_market.asp"-->
<!--#INCLUDE virtual="/include/table_report.asp"-->
<!--#INCLUDE virtual="/include/table_product.asp"-->
<!--#INCLUDE virtual="/include/table_property.asp"-->


<%

content_area="ContentArea"

Set Input=new cInput
period=Input.period
Set Input=nothing

Set u = new cUser

if period=g_CURRENT_PERIOD_ID and u.HasAccessToAllMarketsForReport("itsreport_current")=false then content_area="ErrNoAccess"
if period<>g_CURRENT_PERIOD_ID and u.HasAccessToAllMarketsForReport("itsreport_snapshot")=false then content_area="ErrNoAccess"

'Response.Write "period=" & period & "<br>"
Set u = nothing


Set p=new cPage

Set c = new cConverter
quarter=c.NumToStr("LKP01_PERIOD_NAME", period)
p.Subhead=quarter & " " & Lang("results")
Set c = nothing

Call p.Display(content_area, true)

Set p = nothing

%>


<%
Function ContentArea()
%>

	<center>
	<h3><%=Lang("list_of_graphs_for") & " " & quarter%></h3>

<%
	Call List(0, Lang("standard"))
	Call List(2, Lang("special"))
%>

	</center>
	<br>

<%
End Function


Function List(group, subtitle)
	'Response.Write "period=" & period & "<br>"

	Set list=new cTablePlaintext

	list.DisplayMode=3

	Set t_report=new cTableReport
	Set t_report.TableObj=list
	t_report.CurrPeriodID=period
	t_report.GroupIdFilter=group

	Set t_product=new cTableProduct
	Set t_product.TableObj=list

	Set t_property=new cTableProperty
	Set t_property.TableObj=list
	t_property.CurrPeriodID=period
	t_property.EnableShortNames=false

	report_arr=Split(t_report.Read, ",")
	product_arr=Split(t_product.Read, ",")

	list.DisplayMode=2

	s_list="<table border=""1"" style=""width:75%; margin-bottom: 0px;"">"

	for report=0 to ubound(report_arr) ' report loop

		r=Split(report_arr(report), "|")
		'Response.Write "r(0)=" & r(0) & "<br>"

		chart_id=t_report.GetChartID(r(0))
		if chart_id<0 then chart_id=9999
		t_property.ChartID=chart_id

		s_list=s_list & "<tr><th colspan=""2"" class=""head"">" & Server.HtmlEncode(Ucase(r(1))) & "</th></tr>"

		for product=0 to ubound(product_arr) ' product loop

			'Response.Write "report=" & report & " or " & report_arr(report) & ", product=" & product & " or " & product_arr(product) & "<br>"

			p=Split(product_arr(product), "|")
			'Response.Write "p(0)=" & p(0) & "<br>"

			t_property.ProductID=p(0)
			s_property=t_property.Read

			if len(s_property)>0 then

				rowspan=""
				count_property=Ubound(Split(s_property, ","))+1
				if count_property>1 then rowspan=" rowspan=""" & count_property & """"

				s_property=t_property.ReadNames(s_property, 1)
				s_property="<td align=""left"">" & Server.HtmlEncode(s_property) & "</td>"
				s_property=Replace(s_property, " &#8226; ", "</td></tr><tr><td align=""left"">")
				s_list=s_list & "<tr><th" & rowspan & ">" & Server.HtmlEncode(p(1)) & "</th>" & s_property & "</tr>"

			end if

		next

	next

	s_list=s_list & "</table>"

	Set t_report=nothing
	Set t_product=nothing
	Set t_property=nothing

	Set list=nothing

%>

<!--<h4><%=subtitle%></h4>-->

<%=s_list%>

<%
End Function
%>

<%
Function ErrNoAccess
	Response.Write GetInfoBox(Lang("error_c"), Lang("back"))
End Function
%>
