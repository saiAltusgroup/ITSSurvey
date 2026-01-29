<%

' DK, Apr 01, 2008 - version 1.0

Set gLangDict = CreateObject("Scripting.Dictionary")

Function Lang(item)
	Lang=gLangDict.Item(item)
	if len(Lang)<1 then Lang="[" & item & "]"
	'Lang="*" & Lang
End Function


Function LangField(string)

	i=1
	do
		j1=instr(i, string, "<@L_")
		'Response.Write "j1=" & j1 & "<BR>"
		if j1<=0 then exit do

		j2=instr(j1+1, string, ">")
		'Response.Write "j2=" & j2 & "<BR>"
		if j2<=j1 then exit do

		item=mid(string, j1+2, j2-j1-2)
		item=lang(item)
		'Response.Write "item=" & item & "<BR>"

		string=left(string, j1-1) & item & mid(string, j2+1)

		i=j1+len(item)
	loop

	LangField=string

End Function


host = Request.ServerVariables("SCRIPT_NAME")

if instr(host, "/fre/")>0 then
	gLang="fre"
%>
	<!--#INCLUDE file="language_fre.asp"-->
<%
else
	gLang="eng"
%>
	<!--#INCLUDE file="language_eng.asp"-->
<%
end if
%>
