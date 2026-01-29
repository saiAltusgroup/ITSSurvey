<!--#INCLUDE virtual="/lib/asp/adovbs.asp"-->
<!--#INCLUDE virtual="/sys_include/db.asp"-->
<!--#INCLUDE virtual="/sys_include/globals.asp"-->
<!--#INCLUDE virtual="/include/language.asp"-->
<!--#INCLUDE virtual="/include/user.asp"-->
<!--#INCLUDE virtual="/include/input.asp"-->
<!--#INCLUDE virtual="/include/page.asp"-->
<!--#INCLUDE virtual="/include/converter.asp"-->

<%
Set Input=new cInput
action=Input.action
Set Input=nothing

Set p=new cPage
p.Title=Lang("title_about")
p.Protect=false
p.Include_about=true


Select Case action
	Case "definitions"
		Call p.Display("ContentDefinitions", true)
	Case "copyright"
		Call p.Display("ContentCopyright", true)
	Case "marketdirectory"
		Call p.Display("ContentMarketDirectory", true)
	Case "researchmethodology"
		Call p.Display("ContentReasearchMethodology", true)
	Case "overview"
		Call p.Display("ContentOverview", true)
	Case Else
		Call p.Display("ContentAbout", true)
End Select

Set p = nothing

%>

<%
Function ContentDefinitions
%>
<div id="About">
<!--#include file="about/definitions_benchmark.html"-->
</div><!--/About-->
<%
End Function

Function ContentCopyright
%>
<div id="About">
<!--#include file="about/copyright.html"-->
</div><!--/About-->
<%
End Function

Function ContentAbout
%>
<div id="About">
<!--#include file="about/index.html"-->
</div><!--/About-->
<%
End Function

Function ContentMarketDirectory
%>
<div id="About">
<!--#include file="about/marketdirectory_districtnodes.html"-->
</div><!--/About-->
<%
End Function

Function ContentReasearchMethodology
%>
<div id="About">
<!--#include file="about/researchmethodology.html"-->
</div><!--/About-->
<%
End Function

Function ContentOverview
%>
<div id="About">
<!--#include file="about/overview.html"-->
</div><!--/About-->
<%
End Function
%>