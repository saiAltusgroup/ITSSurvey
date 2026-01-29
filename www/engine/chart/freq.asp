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
<!--#INCLUDE virtual="/include/table_frequency.asp"-->


<%

Set p=new cPage
Call p.Display("ContentArea", true)
Set p = nothing

%>


<%
Function ContentArea()
%>

	<center>
	<h3><%=Lang("menu_qtrfrequency")%></h3>

<%
	Call List()
%>

	</center>
	<br>

<%
End Function


Function List()
	Set t_freq=new cTableFrequency
	Response.Write t_freq.Read
	Set t_freq=nothing
End Function
%>
