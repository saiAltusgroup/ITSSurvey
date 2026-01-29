<!--#INCLUDE virtual="/lib/asp/adovbs.asp"-->
<!--#INCLUDE virtual="/sys_include/globals.asp"-->
<!--#INCLUDE virtual="/sys_include/db.asp"-->
<!--#INCLUDE virtual="/include/language.asp"-->
<!--#INCLUDE virtual="/include/user.asp"-->
<!--#INCLUDE virtual="/include/input.asp"-->
<!--#INCLUDE virtual="/include/page.asp"-->
<!--#INCLUDE virtual="/include/converter.asp"-->

<%

content_area="ContentArea"

Set Input=new cInput
period=Input.period
step=Input.step
market=Input.market
report=Input.report
product=Input.product
property=Input.property
Set Input=nothing

Set u = new cUser

'Response.Write "UserId=" & u.UserId & "<br>"
'Response.Write "HasAccessToAllMarketsForReport(itsreport_hist)=" & u.HasAccessToAllMarketsForReport("itsreport_hist") & "<br>"
'Response.Write "HasAccessToAllMarketsForReport(itsreport_current)=" & u.HasAccessToAllMarketsForReport("itsreport_current") & "<br>"
'Response.Write "HasAccessToAllMarketsForReport(itsreport_snapshot)=" & u.HasAccessToAllMarketsForReport("itsreport_snapshot") & "<br>"

if u.HasAccessToAllMarketsForReport("itsreport_snapshot")=false then content_area="ErrNoAccess"
Set u = nothing


Set p=new cPage
Set c = new cConverter
p.Subhead = Lang("historical_snap")
Set c = nothing

Call p.Display(content_area, true)
Set p = nothing

%>


<%
Function ContentArea
%>
<center>
<h3><%=Lang("hist_qtr_snapshot")%></h3>

<table border="1" width="75%">


<%
for y=Year(Date) to g_MIN_YEAR step -1
%>

<tr><th colspan="4"><%=y%></th></tr>

<tr>

<%
	for q=1 to 4

		period=4*(y-g_MIN_YEAR)+q+(g_MIN_PERIOD_ID-1)
%>
		<td align="center">
		<% if period>=g_MIN_PERIOD_ID and period<=g_MAX_PERIOD_ID then %><a href="default.asp?period=<%=period%>">Q<%=q & " " & y & " " & Lang("results")%></a><% end if %>
		&nbsp;</td>
<%
	next
%>

</tr>
<%
next
%>


</table>
</center>
<br>

<%
End Function
%>

<%
Function ErrNoAccess
	Response.Write GetInfoBox(Lang("error_c"), Lang("back"))
End Function
%>
