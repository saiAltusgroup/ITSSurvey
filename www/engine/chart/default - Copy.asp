<!--#INCLUDE virtual="/lib/asp/adovbs.asp"-->
<!--#INCLUDE virtual="/sys_include/globals.asp"-->
<!--#INCLUDE virtual="/sys_include/db.asp"-->
<!--#INCLUDE virtual="/include/language.asp"-->
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

Set Input=new cInput
period=Input.period
step=Input.step
market=Input.market
report=Input.report
product=Input.product
property=Input.property
Set Input=nothing


Set p=new cPage
'p.Protect=false
p.Include_its=false

Set c = new cConverter
p.Subhead=c.NumToStr("LKP01_PERIOD_NAME", period) & " " & Lang("results")
if period<>g_CURRENT_PERIOD_ID then p.Subhead = p.Subhead & " - " & Lang("historical")
Set c = nothing

if step=0 then

	p.Include_cascade=true
	p.BodyAttribs="onload=""updateStep1("&period&",0,'"&Lang("lang")&"');updateStep1("&period&",1,'"&Lang("lang")&"');"""

	Call p.Display("ContentArea_Ajax", true)
	'Call p.Display("ContentArea_Step_" & step, true)
else
	Call p.Display("ContentArea_Step_" & step, true)
end if

Set p = nothing

%>


<%

Function JSEncode(str_in)
	str=str_in
	str=Replace(str, """", "\""")
	str=Replace(str, "/", "\/")
	JSEncode=str
End Function

Function ContentArea_Ajax
	ContentArea_Ajax = ContentArea(-1)
End Function

Function ContentArea_Step_0
	ContentArea_Step_0 = ContentArea(0)
End Function

Function ContentArea_Step_1
	ContentArea_Step_1 = ContentArea(1)
End Function


Function ContentArea(step)

	random_quote=GetRandomQuote(period)

%>

<!--<p>HasAccessToAllMarketsForReport(itsreport_current)=<%=p.UserObj.HasAccessToAllMarketsForReport("itsreport_current")%></p>-->
<!--<p>HasAccessToAllMarketsForReport(itsreport_hist)=<%=p.UserObj.HasAccessToAllMarketsForReport("itsreport_hist")%></p>-->
<!--<p>HasAccessToAllMarketsForReport(itsreport_snapshot)=<%=p.UserObj.HasAccessToAllMarketsForReport("itsreport_snapshot")%></p>-->

<br><br>

<% 'if period=g_CURRENT_PERIOD_ID then %>

<%
imgs=Array("100","200","300","400", "101","201","301","401", "102","202","302","402", "103","203","303","403", "104","204","304","404", "105","205","305","405", "106","206","306","406", "107","207","307","407", "108", "208", "308", "408", "109", "209", "309", "409", "110", "210", "310", "410", "111", "211", "311", "411", "112", "212", "312", "412", "113", "213", "313", "413", "114", "214", "314", "414", "115", "215","315","415","116","216","316","416","117","217","317","417","118","218","318","418","119","219","319")
i=period-g_FIRST_PERIOD_ID
if i>=0 and i<=ubound(imgs) then img=imgs(i)
if Lang("lang")="fr-ca" and period>=107 then img=img&"-FR"
%>

<a href="result.asp?report=X100&amp;period=<%=period%>&amp;step=2"><img src="/images/IT<%=img%>.gif" alt="<%=Lang("alt_0")%>" align="right" hspace="5" vspace="5" border="0" style="padding-left:1em;padding-bottom:1em;" ></a>

<% 'end if %>


<p class="intro">
<%
if period=150 then
	Response.Write Lang("intro_0_45")
elseif period=149 then
	Response.Write Lang("intro_0_44")	
elseif period=148 then
	Response.Write Lang("intro_0_43")
elseif period=147 then
	Response.Write Lang("intro_0_42")
elseif period=146 then
	Response.Write Lang("intro_0_41")
elseif period=145 then
	Response.Write Lang("intro_0_40")
elseif period=144 then
	Response.Write Lang("intro_0_39")
elseif period=143 then
	Response.Write Lang("intro_0_38")
elseif period=142 then
	Response.Write Lang("intro_0_37")
elseif period=141 then
   	Response.Write Lang("intro_0_36")
elseif period=140 then
	Response.Write Lang("intro_0_35")
elseif period=139 then
    Response.Write Lang("intro_0_34")
elseif period=138 then
	Response.Write Lang("intro_0_33")
elseif period=137 then
	Response.Write Lang("intro_0_32")
elseif period=135 then
	Response.Write Lang("intro_0_31")
elseif period=133 then
	Response.Write Lang("intro_0_30")
elseif period=132 then
	Response.Write Lang("intro_0_29")
elseif period=131 then
	Response.Write Lang("intro_0_28")
elseif period=130 then
	Response.Write Lang("intro_0_27")
elseif period=129 then
	Response.Write Lang("intro_0_26")
elseif period=128 then
	Response.Write Lang("intro_0_25")
elseif period=127 then
	Response.Write Lang("intro_0_24")
elseif period=126 then
	Response.Write Lang("intro_0_23")
elseif period=125 then
	Response.Write Lang("intro_0_22")
elseif period=124 then
	Response.Write Lang("intro_0_21")
elseif period=123 then
	Response.Write Lang("intro_0_20")
elseif period=122 then
	Response.Write Lang("intro_0_19")
elseif period=121 then
	Response.Write Lang("intro_0_18")
elseif period=120 then
	Response.Write Lang("intro_0_17")
elseif period=119 then
	Response.Write Lang("intro_0_16")
elseif period=118 then
	Response.Write Lang("intro_0_15")
elseif period=117 then
	Response.Write Lang("intro_0_14")
elseif period=116 then
	Response.Write Lang("intro_0_13")
elseif period=115 then
	Response.Write Lang("intro_0_12")
elseif period=114 then
	Response.Write Lang("intro_0_11")
elseif period=113 then
	Response.Write Lang("intro_0_10")
elseif period=112 then
	Response.Write Lang("intro_0_9")
elseif period=111 then
	Response.Write Lang("intro_0_8")
elseif period=110 then
	Response.Write Lang("intro_0_7")
elseif period=109 then
	Response.Write Lang("intro_0_6")
elseif period=108 then
	Response.Write Lang("intro_0_5")
elseif period=107 then
	Response.Write Lang("intro_0_4")
elseif period=106 then
	Response.Write Lang("intro_0_3")
elseif period=105 then
	Response.Write Lang("intro_0_2")
elseif period=104 then
	Response.Write Lang("intro_0_1")
else
	Response.Write Lang("intro_0_0")
end if


if Lang("lang")="fr-ca" then
	random_quote_0="&laquo;&nbsp;" & random_quote(0) & "&nbsp;&raquo;"
else
	random_quote_0="&quot;" & random_quote(0) & "&quot;"
end if

%>
</p>

<% if ubound(random_quote)=1 then %>
<blockquote>
<em><strong>
<span style="color:#008833;"><%=random_quote_0%></span><br>
<small>&mdash; <%=random_quote(1)%></small>
</strong></em>
</blockquote>
<% end if %>

<p class="intro">
<%=Lang("intro_1")%>
</p>


<div style="clear:both;"></div>

<!--<h3 style="clear:both;"><%=s_step_title%></h3>-->

<%
	Set t=new cTableReport
	groupid=t.GetGroupId(report)
	Set t=nothing

	if groupid=-1 then
%>
	<table border="0" class="itslayout"><tr valign="top"><td valign="top" class="formleft" width="50%">
<%
		Call FormArea(0, step)
		Call FormArea(2, 2)
%>
		</td><td valign="top" class="formright" width="50%">
<%
		Call FormArea(1, step)
%>
		</td></tr></table>
<%
	else
		Call FormArea(groupid, step)
	end if

End Function


Function FormArea(groupid, step_in)

step=step_in

' step = -1 : Current & Historical Reports, Ajax Version
' step = 0 : Current & Historical Reports, Step 0, plain HTML version
' step = 1 : Current & Historical Reports, Step 1, plain HTML version
' step = 2 : Special Reports, Plain HTML version

if groupid=2 then
	s_title=Lang("form_a_title")
	s_intro=Lang("form_a_instr")
	s_cssclass="lookupform"
elseif groupid=1 then
	s_title=Lang("form_b_title")
	s_intro=Lang("form_b_instr")
	s_cssclass="searchform"
else
	s_title=Lang("form_c_title")
	s_intro=Lang("form_c_instr")
	s_cssclass="lookupform"
end if



ajax=false
s_submit_page=Request.ServerVariables("SCRIPT_NAME")
s_submit_step="1"
s_submit_text=Lang("next")

if step=-1 then
	ajax=true
	step=0
end if



extra_attribs1="onchange=""updateStep1("&period&","&groupid&",'"&Lang("lang")&"');"" onfocus=""updateStep1("&period&","&groupid&",'"&Lang("lang")&"');"""
extra_attribs2=""
if groupid=1 then extra_attribs2="onchange=""updateMarket();"" onfocus=""updateMarket();"""

'option_0=""
'option_0="<option value=""""></option>"
option_0="<option value="""">" & Lang("select") & "</option>"

s_report="<select name=""report_" & groupid & """ id=""report_" & groupid & """>" & option_0 & "</select>"
s_product="<select name=""product_" & groupid & """ id=""product_" & groupid & """>" & option_0 & "</select>"
s_property="<select name=""property_" & groupid & """ id=""property_" & groupid & """ " & extra_attribs2 & ">" & option_0 & "</select>"
s_market="<select name=""market_" & groupid & """ id=""market_" & groupid & """>" & option_0 & "</select>"



if step=0 then

	Set listbox=new cTableListbox
	listbox.FullDisplay=false

	Set t=new cTableReport
	t.CurrPeriodID=period
	t.DeactivateExceptionsByPeriod(period)
	t.GroupIdFilter=groupid
	listbox.Name="report_" & groupid
	if ajax then listbox.ExtraAttribs=extra_attribs1
	Set t.TableObj=listbox
	if groupid=1 then t.CheckExistsInPeriod=false
	s_report=t.Read
	Set t=nothing

	Set t=new cTableProduct
	listbox.Name="product_" & groupid
	if ajax then listbox.ExtraAttribs=extra_attribs1
	Set t.TableObj=listbox
	s_product=t.Read
	if groupid=0 then s_product=listbox.PrependHack(s_product, "0", Lang("all_products"), true)
	Set t = nothing

	Set listbox=nothing

	if ajax then
		s_step_title=Lang("step_title_ajax")
		s_submit_text=Lang("display_report")
	else
		s_step_title=Lang("step_title_1b")
	end if

elseif step=1 then

	Set tr=new cTableReport
	tr.CurrPeriodID=period
	tr.DeactivateExceptionsByPeriod(period)
	s_property=PropertySelector_HTML(period, tr.GetPropertySelectorType(report), product, tr.GetChartID(report), groupid=1)
	s_market=MarketSelector_HTML(period, tr.GetMarketSelectorType(report), tr.GetChartID(report), groupid=1)
	Set tr=nothing

	s_step_title=Lang("step_title_2a")

	s_submit_page="result.asp"
	s_submit_step="2"
	s_submit_text=Lang("display_report")

elseif step=2 then

	Set listbox=new cTableListbox
	listbox.FullDisplay=false

	Set t=new cTableReport
	t.CurrPeriodID=period
	t.DeactivateExceptionsByPeriod(period)
	t.GroupIdFilter=groupid
	listbox.Name="report_" & groupid
	Set t.TableObj=listbox
	s_report=t.Read
	Set t=nothing

	Set listbox=nothing

	s_step_title=Lang("step_title_1a")

	s_submit_page="result.asp"
	s_submit_step="2"
	s_submit_text=Lang("display_report")

end if


%>

<form action="<%=s_submit_page%>" method="post" class="<%=s_cssclass%>" style="width:99%;" onsubmit="return handleForm(this);">

<fieldset>
<legend><%=s_title%></legend>

<p>
<%=s_intro%>
</p>

<% if step=0 then %>

	<% if Len(s_report)>0 then %><label for="report_<%=groupid%>"><%=Lang("report_label")%></label> <%=s_report%><% end if %>
	<% if Len(s_product)>0 then %><label for="product_<%=groupid%>"><%=Lang("product_label")%></label> <%=s_product%><% end if %>

		<% if ajax then %>

			<script type="text/javascript">
			<!-- to hide script contents from old browsers

			document.write("<label for=\"property_<%=groupid%>\"><%=Lang("property_label")%><\/label> <%=JSEncode(s_property)%>");
			document.write("<label for=\"market_<%=groupid%>\"><%=Lang("market_label")%><\/label> <%=JSEncode(s_market)%>");

			// end hiding contents from old browsers -->
			</script>

		<% end if %>

<% elseif step=1 then %>


	<input type="hidden" name="report_<%=groupid%>" value="<%=report%>">
	<input type="hidden" name="product_<%=groupid%>" value="<%=product%>">


	<% if Len(s_property)>0 then %><label for="property"><%=Lang("property_label")%></label> <%=s_property%><% end if %>
	<% if Len(s_market)>0 then %><label for="market"><%=Lang("market_label")%></label> <%=s_market%><% end if %>

<% elseif step=2 then %>

	<label for="report_<%=groupid%>"><%=Lang("report_label")%></label> <%=s_report%>

<% end if %>

<% if groupid=1 then %>

<input type="radio" name="histlen" id="histlen_1" value="0" checked><label for="histlen_1" style="display: inline;"><%=Lang("history_all")%></label><br>
<input type="radio" name="histlen" id="histlen_2" value="5"><label for="histlen_2" style="display: inline;"><%=Lang("history_top5")%></label><br>

<% end if %>



<input type="hidden" name="period" value="<%=period%>">
<input type="hidden" name="step" value="<%=s_submit_step%>">
<br><input type="submit" value="<%=s_submit_text%>">



</fieldset>

<% if step=0 then %>

<noscript>
<fieldset>
<legend>!</legend>
<p>
<%=Lang("info_javascript")%>
</p>
</fieldset>
</noscript>

<% end if %>

</form>


<%
End Function
%>


<%
Function GetRandomQuote(period)

	'Response.Write "period=" & period

	if period=107 then
		n=13
		ReDim q(n)

		q(0)=Array(Lang("quote_a0_3"), Lang("quote_b0_3"))
		q(1)=Array(Lang("quote_a1_3"), Lang("quote_b1_3"))
		q(2)=Array(Lang("quote_a2_3"), Lang("quote_b2_3"))
		q(3)=Array(Lang("quote_a3_3"), Lang("quote_b3_3"))
		q(4)=Array(Lang("quote_a4_3"), Lang("quote_b4_3"))
		q(5)=Array(Lang("quote_a5_3"), Lang("quote_b5_3"))
		q(6)=Array(Lang("quote_a6_3"), Lang("quote_b6_3"))
		q(7)=Array(Lang("quote_a7_3"), Lang("quote_b7_3"))
		q(8)=Array(Lang("quote_a8_3"), Lang("quote_b8_3"))
		q(9)=Array(Lang("quote_a9_3"), Lang("quote_b9_3"))
		q(10)=Array(Lang("quote_a10_3"), Lang("quote_b10_3"))
		q(11)=Array(Lang("quote_a11_3"), Lang("quote_b11_3"))
		q(12)=Array(Lang("quote_a12_3"), Lang("quote_b12_3"))
		q(13)=Array(Lang("quote_a13_3"), Lang("quote_b13_3"))

	elseif period=106 then
		n=5
		ReDim q(n)

		q(0)=Array(Lang("quote_a0_2"), Lang("quote_b0_2"))
		q(1)=Array(Lang("quote_a1_2"), Lang("quote_b1_2"))
		q(2)=Array(Lang("quote_a2_2"), Lang("quote_b2_2"))
		q(3)=Array(Lang("quote_a3_2"), Lang("quote_b3_2"))
		q(4)=Array(Lang("quote_a4_2"), Lang("quote_b4_2"))
		q(5)=Array(Lang("quote_a5_2"), Lang("quote_b5_2"))

	elseif period=105 then
		n=4
		ReDim q(n)

		q(0)=Array(Lang("quote_a0_1"), Lang("quote_b0_1"))
		q(1)=Array(Lang("quote_a1_1"), Lang("quote_b1_1"))
		q(2)=Array(Lang("quote_a2_1"), Lang("quote_b2_1"))
		q(3)=Array(Lang("quote_a3_1"), Lang("quote_b3_1"))
		q(4)=Array(Lang("quote_a4_1"), Lang("quote_b4_1"))

	elseif period=104 then
		n=7
		ReDim q(n)

		q(0)=Array(Lang("quote_a0_0"), Lang("quote_b0_0"))
		q(1)=Array(Lang("quote_a1_0"), Lang("quote_b1_0"))
		q(2)=Array(Lang("quote_a2_0"), Lang("quote_b2_0"))
		q(3)=Array(Lang("quote_a3_0"), Lang("quote_b3_0"))
		q(4)=Array(Lang("quote_a4_0"), Lang("quote_b4_0"))
		q(5)=Array(Lang("quote_a5_0"), Lang("quote_b5_0"))
		q(6)=Array(Lang("quote_a6_0"), Lang("quote_b6_0"))
		q(7)=Array(Lang("quote_a7_0"), Lang("quote_b7_0"))
	else
		n=0
		ReDim q(n)

		q(0)=Array("")

	end if

	Randomize
	i=Int((n+1)*Rnd())
	GetRandomQuote=q(i)

	'Response.Write i

End Function

%>
