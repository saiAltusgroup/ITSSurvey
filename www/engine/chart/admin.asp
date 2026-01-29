<!--#INCLUDE virtual="/lib/asp/adovbs.asp"-->
<!--#INCLUDE virtual="/sys_include/globals.asp"-->
<!--#INCLUDE virtual="/sys_include/db.asp"-->
<!--#INCLUDE virtual="/include/language.asp"-->
<!--#INCLUDE virtual="/include/user.asp"-->
<!--#INCLUDE virtual="/include/input.asp"-->
<!--#INCLUDE virtual="/include/page.asp"-->
<!--#INCLUDE virtual="/include/converter.asp"-->

<%

Set p=new cPage
p.Title=Lang("title_superuser")
p.Subhead=p.Title
if p.UserObj.AdminLevel=100 then
	Call p.Display("ContentArea", true)
else
	Call p.Display("NoAccess", true)
end if
Set p = nothing

%>

<%
Function ContentArea
%>

<p><strong>WARNING: DO NOT USE THESE FUNCTIONS UNLESS YOU KNOW WHAT YOU ARE DOING.</strong></p>

<ul>
<li><a href="test.asp">Chart Test Page</a></li>
<li><a href="../import/">Data Importer and Filter</a></li>
<li><a href="../import/chart_column.asp">Column Config Replicator</a></li>
</ul>

<%
End Function
%>


<%
Function NoAccess
%>
<h3>You are not authorized to access this page</h3>
<%
End Function
%>
