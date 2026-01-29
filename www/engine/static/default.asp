<!--#INCLUDE virtual="/lib/asp/adovbs.asp"-->
<!--#INCLUDE virtual="/sys_include/globals.asp"-->
<!--#INCLUDE virtual="/sys_include/db.asp"-->
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
p.Protect=false
p.Include_about=true

Select Case action
	Case "2008_Q1"
		p.Title=Lang("title_outlook_2008_1")
		Call p.Display("ContentOutlook_2008_Q1", true)
	Case "2008_Q2"
		p.Title=Lang("title_outlook_2008_2")
		Call p.Display("ContentOutlook_2008_Q2", true)
	Case "2008_Q3"
		p.Title=Lang("title_outlook_2008_3")
		Call p.Display("ContentOutlook_2008_Q3", true)
	Case "2008_Q4"
		p.Title=Lang("title_outlook_2008_4")
		Call p.Display("ContentOutlook_2008_Q4", true)
	Case "2009_Q1"
		p.Title=Lang("title_outlook_2009_1")
		Call p.Display("ContentOutlook_2009_Q1", true)
	Case "2009_Q2"
		p.Title=Lang("title_outlook_2009_2")
		Call p.Display("ContentOutlook_2009_Q2", true)
	Case "2009_Q3"
		p.Title=Lang("title_outlook_2009_3")
		Call p.Display("ContentOutlook_2009_Q3", true)
	Case "2009_Q4"
		p.Title=Lang("title_outlook_2009_4")
		Call p.Display("ContentOutlook_2009_Q4", true)
	Case "2010_Q1"
		p.Title=Lang("title_outlook_2010_1")
		Call p.Display("ContentOutlook_2010_Q1", true)
	Case "2010_Q2"
		p.Title=Lang("title_outlook_2010_2")
		Call p.Display("ContentOutlook_2010_Q2", true)
	Case "2010_Q3"
		p.Title=Lang("title_outlook_2010_3")
		Call p.Display("ContentOutlook_2010_Q3", true)
	Case "2010_Q4"
		p.Title=Lang("title_outlook_2010_4")
		Call p.Display("ContentOutlook_2010_Q4", true)
	Case "2011_Q1"
		p.Title=Lang("title_outlook_2011_1")
		Call p.Display("ContentOutlook_2011_Q1", true)
	Case "2011_Q2"
		p.Title=Lang("title_outlook_2011_2")
		Call p.Display("ContentOutlook_2011_Q2", true)
	Case "2011_Q3"
		p.Title=Lang("title_outlook_2011_3")
		Call p.Display("ContentOutlook_2011_Q3", true)
	Case "2011_Q4"
		p.Title=Lang("title_outlook_2011_4")
		Call p.Display("ContentOutlook_2011_Q4", true)
	Case "2012_Q1"
		p.Title=Lang("title_outlook_2012_1")
		Call p.Display("ContentOutlook_2012_Q1", true)
	Case "2012_Q2"
		p.Title=Lang("title_outlook_2012_2")
		Call p.Display("ContentOutlook_2012_Q2", true)
End Select

Set p = nothing

%>

<%
Function ContentOutlook_2008_Q1
%>
<div id="About">
<% If Lang("lang")="fr-ca" Then %>
	<!--#include file="static_fr/Outlook_2008_Q1.html"-->
<% Else %>
	<!--#include file="static_en/Outlook_2008_Q1.html"-->
<% End If %>
</div>
<%
End Function

Function ContentOutlook_2008_Q2
%>
<div id="About">
<% If Lang("lang")="fr-ca" Then %>
	<!--#include file="static_fr/Outlook_2008_Q2.html"-->
<% Else %>
	<!--#include file="static_en/Outlook_2008_Q2.html"-->
<% End If %>
</div>
<%
End Function

Function ContentOutlook_2008_Q3
%>
<div id="About">
<% If Lang("lang")="fr-ca" Then %>
	<!--#include file="static_fr/Outlook_2008_Q3.html"-->
<% Else %>
	<!--#include file="static_en/Outlook_2008_Q3.html"-->
<% End If %>
</div>
<%
End Function

Function ContentOutlook_2008_Q4
%>
<div id="About">
<% If Lang("lang")="fr-ca" Then %>
	<!--#include file="static_fr/Outlook_2008_Q4.html"-->
<% Else %>
	<!--#include file="static_en/Outlook_2008_Q4.html"-->
<% End If %>
</div>
<%
End Function

Function ContentOutlook_2009_Q1
%>
<div id="About">
<% If Lang("lang")="fr-ca" Then %>
	<!--#include file="static_fr/Outlook_2009_Q1.html"-->
<% Else %>
	<!--#include file="static_en/Outlook_2009_Q1.html"-->
<% End If %>
</div>
<%
End Function

Function ContentOutlook_2009_Q2
%>
<div id="About">
<% If Lang("lang")="fr-ca" Then %>
	<!--#include file="static_fr/Outlook_2009_Q2.html"-->
<% Else %>
	<!--#include file="static_en/Outlook_2009_Q2.html"-->
<% End If %>
</div>
<%
End Function

Function ContentOutlook_2009_Q3
%>
<div id="About">
<% If Lang("lang")="fr-ca" Then %>
	<!--#include file="static_fr/Outlook_2009_Q3.html"-->
<% Else %>
	<!--#include file="static_en/Outlook_2009_Q3.html"-->
<% End If %>
</div>
<%
End Function

Function ContentOutlook_2009_Q4
%>
<div id="About">
<% If Lang("lang")="fr-ca" Then %>
	<!--#include file="static_fr/Outlook_2009_Q4.html"-->
<% Else %>
	<!--#include file="static_en/Outlook_2009_Q4.html"-->
<% End If %>
</div>
<%
End Function

Function ContentOutlook_2010_Q1
%>
<div id="About">
<% If Lang("lang")="fr-ca" Then %>
	<!--#include file="static_fr/Outlook_2010_Q1.html"-->
<% Else %>
	<!--#include file="static_en/Outlook_2010_Q1.html"-->
<% End If %>
</div>
<%
End Function

Function ContentOutlook_2010_Q2
%>
<div id="About">
<% If Lang("lang")="fr-ca" Then %>
	<!--#include file="static_fr/Outlook_2010_Q2.html"-->
<% Else %>
	<!--#include file="static_en/Outlook_2010_Q2.html"-->
<% End If %>
</div>
<%
End Function

Function ContentOutlook_2010_Q3
%>
<div id="About">
<% If Lang("lang")="fr-ca" Then %>
	<!--#include file="static_fr/Outlook_2010_Q3.html"-->
<% Else %>
	<!--#include file="static_en/Outlook_2010_Q3.html"-->
<% End If %>
</div>
<%
End Function

Function ContentOutlook_2010_Q4
%>
<div id="About">
<% If Lang("lang")="fr-ca" Then %>
	<!--#include file="static_fr/Outlook_2010_Q4.html"-->
<% Else %>
	<!--#include file="static_en/Outlook_2010_Q4.html"-->
<% End If %>
</div>
<%
End Function

Function ContentOutlook_2011_Q1
%>
<div id="About">
<% If Lang("lang")="fr-ca" Then %>
	<!--#include file="static_fr/Outlook_2011_Q1.html"-->
<% Else %>
	<!--#include file="static_en/Outlook_2011_Q1.html"-->
<% End If %>
</div>
<%
End Function

Function ContentOutlook_2011_Q2
%>
<div id="About">
<% If Lang("lang")="fr-ca" Then %>
	<!--#include file="static_fr/Outlook_2011_Q2.html"-->
<% Else %>
	<!--#include file="static_en/Outlook_2011_Q2.html"-->
<% End If %>
</div>
<%
End Function

Function ContentOutlook_2011_Q3
%>
<div id="About">
<% If Lang("lang")="fr-ca" Then %>
	<!--#include file="static_fr/Outlook_2011_Q3.html"-->
<% Else %>
	<!--#include file="static_en/Outlook_2011_Q3.html"-->
<% End If %>
</div>
<%
End Function

Function ContentOutlook_2011_Q4
%>
<div id="About">
<% If Lang("lang")="fr-ca" Then %>
	<!--#include file="static_fr/Outlook_2011_Q4.html"-->
<% Else %>
	<!--#include file="static_en/Outlook_2011_Q4.html"-->
<% End If %>
</div>
<%
End Function

Function ContentOutlook_2012_Q1
%>
<div id="About">
<% If Lang("lang")="fr-ca" Then %>
	<!--#include file="static_fr/Outlook_2012_Q1.html"-->
<% Else %>
	<!--#include file="static_en/Outlook_2012_Q1.html"-->
<% End If %>
</div>
<%
End Function

Function ContentOutlook_2012_Q2
%>
<div id="About">
<% If Lang("lang")="fr-ca" Then %>
	<!--#include file="static_fr/Outlook_2012_Q2.html"-->
<% Else %>
	<!--#include file="static_en/Outlook_2012_Q2.html"-->
<% End If %>
</div>
<%
End Function
%>
